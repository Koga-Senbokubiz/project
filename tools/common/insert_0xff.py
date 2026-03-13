# replace_0102_to_ff.py

import sys

def replace_bytes(input_file, output_file):
    with open(input_file, "rb") as f:
        data = f.read()

    # 0x01 0x02 → 0xFF
    replaced = data.replace(b'\x01\x02', b'\xff')

    with open(output_file, "wb") as f:
        f.write(replaced)

    print("変換完了")
    print("input :", input_file)
    print("output:", output_file)


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python replace_0102_to_ff.py input_file output_file")
        sys.exit(1)

    replace_bytes(sys.argv[1], sys.argv[2])