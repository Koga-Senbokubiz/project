from __future__ import annotations

import argparse
import configparser
import logging
from datetime import datetime
from pathlib import Path
from typing import Sequence

from bbord_xml_to_dfax import BBORDXMLtoDFAX, CRLF


DEFAULT_INI_NAME = "bbord_fax.ini"
DEFAULT_SECTION = "dfax"


def build_req_no(now: datetime) -> str:
    return f"VANBB1{now.strftime('%y')}{now.timetuple().tm_yday:03d}{now.hour * 3600 + now.minute * 60 + now.second:05d}"


def build_i0(now: datetime) -> bytes:
    req_no = build_req_no(now)
    return (
        b"I0"
        + b"VAN     "
        + b"BB1     "
        + req_no.encode("ascii")
        + b" " * 55
        + b"03"
        + b"003"
        + b" "
        + b"VANBB1          "
        + b"VANBB11         "
    )


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="BBORD_FAX.php / BBORD_XMLtoFAX.php compatible D-FAX generator"
    )
    parser.add_argument("jobid", help="job id for log output")
    parser.add_argument("input_dir", help="input XML directory")
    parser.add_argument("output_file", help="output DFAXDATA file path")
    parser.add_argument(
        "--ini",
        default=str(Path(__file__).with_name(DEFAULT_INI_NAME)),
        help="ini file path",
    )
    parser.add_argument(
        "--fixed-now",
        help="override current timestamp with YYYY-MM-DD HH:MM:SS",
    )
    return parser.parse_args(argv)


def load_settings(ini_path: str | Path) -> dict[str, str]:
    path = Path(ini_path)
    if not path.exists():
        raise FileNotFoundError(f"ini file not found: {path}")

    config = configparser.ConfigParser()
    read_files = config.read(path, encoding="utf-8-sig")
    if not read_files:
        raise RuntimeError(f"failed to read ini file: {path}")
    if DEFAULT_SECTION not in config:
        raise KeyError(f"section [{DEFAULT_SECTION}] not found in ini: {path}")

    sec = config[DEFAULT_SECTION]
    settings = {
        "log_dir": sec.get("log_dir", "").strip(),
        "log_file_name": sec.get("log_file_name", "").strip(),
        "snd_id": sec.get("snd_id", "").strip(),
        "bb_fax_no": sec.get("bb_fax_no", "").strip(),
        "debug_mode": sec.get("debug_mode", "0").strip(),
        "debug_fax_no": sec.get("debug_fax_no", "").strip(),
        "def_xlsx": sec.get("def_xlsx", "").strip(),
        "log_level": sec.get("log_level", "INFO").strip().upper() or "INFO",
    }

    missing = [
        k
        for k in (
            "log_dir",
            "log_file_name",
            "snd_id",
            "bb_fax_no",
            "debug_mode",
            "debug_fax_no",
        )
        if not settings[k]
    ]
    if missing:
        raise KeyError(
            f"missing required ini keys in [{DEFAULT_SECTION}]: {', '.join(missing)}"
        )

    return settings


def configure_logger(
    jobid: str,
    log_dir: str | Path,
    log_file_name: str,
    log_level: str = "INFO",
) -> logging.Logger:
    base_dir = Path(log_dir)
    job_dir = base_dir / jobid
    job_dir.mkdir(parents=True, exist_ok=True)
    log_file = job_dir / log_file_name

    logger = logging.getLogger("dfax")
    logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
    logger.propagate = False

    # 既存ハンドラを安全に解放
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
        try:
            handler.close()
        except Exception:
            pass

    formatter = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(name)s - %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S"
    )

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)

    return logger


def validate_settings(settings: dict[str, str]) -> None:
    if len(settings["snd_id"]) != 6 or not settings["snd_id"].isdigit():
        raise ValueError(f"HAISIN:'{settings['snd_id']}' is wrong!")
    if len(settings["bb_fax_no"]) != 10 or not settings["bb_fax_no"].isdigit():
        raise ValueError(f"BBFAX:'{settings['bb_fax_no']}' is wrong!")
    if len(settings["debug_fax_no"]) != 10 or not settings["debug_fax_no"].isdigit():
        raise ValueError(f"DEBUG FAX:'{settings['debug_fax_no']}' is wrong!")
    if settings["debug_mode"] not in {"0", "Y"}:
        raise ValueError(
            f"DEBUG MODE:'{settings['debug_mode']}' is wrong! expected 0 or Y"
        )


def run(
    jobid: str,
    input_dir: str | Path,
    output_file: str | Path,
    ini_path: str | Path | None = None,
    fixed_now: str | None = None,
) -> int:
    ini_path = Path(ini_path) if ini_path else Path(__file__).with_name(DEFAULT_INI_NAME)

    settings = load_settings(ini_path)
    validate_settings(settings)
    log_level = "DEBUG" if settings["debug_mode"] == "Y" else "INFO"

    logger = configure_logger(
        jobid,
        settings["log_dir"],
        settings["log_file_name"],
        log_level,
    )
    logger.info("D-FAX変換開始 jobid=%s ini=%s", jobid, ini_path)

    input_dir = Path(input_dir)
    out_file = Path(output_file)
    def_xlsx = settings["def_xlsx"] or None

    try:
        out_file.parent.mkdir(parents=True, exist_ok=True)

        if out_file.exists():
            logger.info("既存出力ファイルを削除: %s", out_file)
            out_file.unlink()

        now = (
            datetime.strptime(fixed_now, "%Y-%m-%d %H:%M:%S")
            if fixed_now
            else datetime.now()
        )

        logger.info("入力フォルダ=%s 出力ファイル=%s", input_dir, out_file)

        if not input_dir.exists() or not input_dir.is_dir():
            raise FileNotFoundError(
                f"input_dir not found or not directory: {input_dir}"
            )

        targets = sorted(
            [p for p in input_dir.iterdir() if p.is_file() and p.stat().st_size > 0]
        )
        logger.info("入力対象ファイル数=%s", len(targets))

        if not targets:
            logger.warning("入力対象ファイルなし")
            logger.info("D-FAX変換終了 jobid=%s", jobid)
            return 0

        with out_file.open("wb") as fp:
            fp.write(build_i0(now) + CRLF)

            for file in targets:
                tran = BBORDXMLtoDFAX(
                    out_fp=fp,
                    bb_fax_no=settings["bb_fax_no"],
                    debug_mode=settings["debug_mode"],
                    debug_fax_no=settings["debug_fax_no"],
                    now=now,
                    def_xlsx=def_xlsx,
                    logger=logger.getChild("tran"),
                )

                logger.info("処理ファイル: %s", file)
                #logger.info("入力ファイル処理開始: %s", file)
                ok, msg = tran.do_tran(file)
                if not ok:
                    raise RuntimeError(msg)
                #logger.info("入力ファイル処理完了: %s", file)

            fp.write(b"S0" + CRLF)
            logger.info("S0出力完了")

        logger.info("D-FAX変換終了 jobid=%s", jobid)
        return 0

    except Exception:
        logger.exception("D-FAX変換異常終了 jobid=%s", jobid)
        raise


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    return run(
        jobid=args.jobid,
        input_dir=args.input_dir,
        output_file=args.output_file,
        ini_path=args.ini,
        fixed_now=args.fixed_now,
    )


if __name__ == "__main__":
    raise SystemExit(main())