#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Step1
顧客XML → 正規化XML
"""

import os
import subprocess
import sys


def main():

    if len(sys.argv) < 3:
        print("usage:")
        print("python order_x2f_step1.py input.xml output.xml")
        sys.exit(1)

    input_xml = sys.argv[1]
    output_xml = sys.argv[2]

    script_dir = os.path.dirname(os.path.abspath(__file__))

    # common フォルダ
    normalize_script = os.path.abspath(
        os.path.join(script_dir, "..", "..", "common", "normalize_order_xml.py")
    )

    cmd = [
        "python",
        normalize_script,
        input_xml,
        "-o",
        output_xml,
        "--strip-namespace",
        "--limit-order",
        "1"
    ]

    print("execute:")
    print(" ".join(cmd))

    subprocess.run(cmd, check=True)

    print("Step1 finished")
    print(f"output : {output_xml}")


if __name__ == "__main__":
    main()
