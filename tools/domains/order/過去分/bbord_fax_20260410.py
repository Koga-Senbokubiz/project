from __future__ import annotations

import argparse
from datetime import datetime
from pathlib import Path
from typing import Sequence

from bbord_xml_to_dfax import BBORDXMLtoDFAX, CRLF


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
    parser = argparse.ArgumentParser(description="BBORD_FAX.php / BBORD_XMLtoFAX.php compatible D-FAX generator")
    parser.add_argument("folder", help="input XML folder")
    parser.add_argument("out_file", help="output DFAXDATA.txt path")
    parser.add_argument("snd_id", help="6 digit sender id")
    parser.add_argument("bb_fax_no", help="10 digit BigBoss fax number")
    parser.add_argument("debug_mode", help="debug mode, usually 0 or Y")
    parser.add_argument("debug_fax_no", help="10 digit debug fax number")
    parser.add_argument("--fixed-now", help="override current timestamp with YYYY-MM-DD HH:MM:SS")
    parser.add_argument("--def-xlsx", help="optional Excel definition path for incremental override")
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    if len(args.snd_id) != 6 or not args.snd_id.isdigit():
        raise SystemExit(f"HAISIN:'{args.snd_id}' is wrong!")
    if len(args.bb_fax_no) != 10 or not args.bb_fax_no.isdigit():
        raise SystemExit(f"BBFAX:'{args.bb_fax_no}' is wrong!")
    if len(args.debug_fax_no) != 10 or not args.debug_fax_no.isdigit():
        raise SystemExit(f"DEBUG FAX:'{args.debug_fax_no}' is wrong!")

    folder = Path(args.folder)
    out_file = Path(args.out_file)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    if out_file.exists():
        out_file.unlink()

    now = datetime.strptime(args.fixed_now, "%Y-%m-%d %H:%M:%S") if args.fixed_now else datetime.now()
    targets = sorted([p for p in folder.iterdir() if p.is_file() and p.stat().st_size > 0])
    if not targets:
        return 0

    with out_file.open("wb") as fp:
        fp.write(build_i0(now) + CRLF)
        tran = BBORDXMLtoDFAX(
            out_fp=fp,
            bb_fax_no=args.bb_fax_no,
            debug_mode=args.debug_mode,
            debug_fax_no=args.debug_fax_no,
            now=now,
            def_xlsx=args.def_xlsx,
        )
        for file in targets:
            ok, msg = tran.do_tran(file)
            if not ok:
                raise RuntimeError(msg)
        fp.write(b"S0" + CRLF)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
