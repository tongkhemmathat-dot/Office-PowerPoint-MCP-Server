#!/usr/bin/env python
"""
Download the AWS service icons (PNG) used by the add_icon / list_icons tools.

Source: https://github.com/awslabs/aws-icons-for-plantuml (64px PNGs, derived from the
official AWS Architecture Icons). The icons remain subject to AWS's trademark and
architecture-icon usage terms: https://aws.amazon.com/architecture/icons/

Usage:
    python scripts/fetch_aws_icons.py [--dest icons/aws]

Icons are flattened into one directory and named after the file (e.g. EC2.png,
ElasticLoadBalancingApplicationLoadBalancer.png). Dark variants (*_Dark.png) are skipped.
"""
import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

REPO = "https://github.com/awslabs/aws-icons-for-plantuml.git"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    root = Path(__file__).resolve().parent.parent
    parser.add_argument("--dest", default=str(root / "icons" / "aws"), help="Output directory")
    dest = Path(parser.parse_args().dest)

    if shutil.which("git") is None:
        print("git is required", file=sys.stderr)
        return 1

    with tempfile.TemporaryDirectory() as tmp:
        clone = Path(tmp) / "repo"
        subprocess.run(["git", "clone", "--depth", "1", "--filter=blob:none", "--sparse", REPO, str(clone)], check=True)
        subprocess.run(["git", "-C", str(clone), "sparse-checkout", "set", "dist"], check=True)

        dest.mkdir(parents=True, exist_ok=True)
        count = 0
        for png in sorted((clone / "dist").rglob("*.png")):
            if png.stem.endswith("_Dark"):
                continue
            # First one wins when two categories share a file name.
            target = dest / png.name
            if not target.exists():
                shutil.copyfile(png, target)
                count += 1
    print(f"Copied {count} icons to {dest}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
