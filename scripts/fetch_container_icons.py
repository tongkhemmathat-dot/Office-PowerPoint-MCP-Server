#!/usr/bin/env python
"""
Download icons for Docker / Kubernetes diagrams into icons/ (used by add_icon / list_icons).

  k8s     Kubernetes community icon set (resources, control plane, infrastructure) from
          https://github.com/kubernetes/community/tree/master/icons  (Apache-2.0).
          Written to icons/k8s/k8s-<name>.png, e.g. k8s-deployment.png, k8s-service.png.
  docker  Docker and related logos from https://github.com/devicons/devicon (MIT), SVG converted
          to PNG with Microsoft Edge / Chrome in headless mode (Windows/macOS/Linux with either
          installed). Written to icons/docker/logo-<name>.png.

Trademarks: the Docker and Kubernetes names and logos belong to their owners; follow their brand
guidelines (https://www.docker.com/legal/trademark-guidelines/, https://www.linuxfoundation.org/legal/trademark-usage).

Usage:
    python scripts/fetch_container_icons.py [--only k8s|docker] [--dest icons]
"""
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image

K8S_REPO = "https://github.com/kubernetes/community.git"
DEVICON_REPO = "https://github.com/devicons/devicon.git"

# short file prefix in the upstream set -> readable name
K8S_NAMES = {
    "c-role": "cluster-role", "crb": "cluster-role-binding", "crd": "custom-resource-definition",
    "cm": "configmap", "deploy": "deployment", "ds": "daemonset", "ep": "endpoints",
    "hpa": "horizontal-pod-autoscaler", "ing": "ingress", "limits": "limit-range",
    "netpol": "network-policy", "ns": "namespace", "psp": "pod-security-policy",
    "pv": "persistent-volume", "pvc": "persistent-volume-claim", "quota": "resource-quota",
    "rb": "role-binding", "rs": "replicaset", "sa": "service-account", "sc": "storage-class",
    "sts": "statefulset", "svc": "service", "vol": "volume",
    "api": "api-server", "c-c-m": "cloud-controller-manager", "c-m": "controller-manager",
    "k-proxy": "kube-proxy", "sched": "scheduler",
}
# (folder, variant) to read; control plane components only exist as labeled icons
K8S_SOURCES = [("resources", "unlabeled"), ("infrastructure_components", "unlabeled"),
               ("control_plane_components", "labeled")]

DOCKER_LOGOS = [  # (devicon folder, svg file, output name)
    ("docker", "docker-original.svg", "docker"),
    ("docker", "docker-original-wordmark.svg", "docker-wordmark"),
    ("kubernetes", "kubernetes-plain.svg", "kubernetes"),
    ("podman", "podman-original.svg", "podman"),
    ("helm", "helm-original.svg", "helm"),
]


def run(cmd, **kw):
    subprocess.run(cmd, check=True, **kw)


def sparse_clone(repo, path, dest):
    run(["git", "clone", "-q", "--depth", "1", "--filter=blob:none", "--sparse", repo, str(dest)])
    run(["git", "-C", str(dest), "sparse-checkout", "set", path])


def biggest(files):
    """Largest readable image (a few upstream files, e.g. etcd-.png, are not valid PNGs)."""
    sizes = {}
    for f in files:
        try:
            with Image.open(f) as im:
                sizes[f] = im.size[0] * im.size[1]
        except Exception:
            continue
    if not sizes:
        raise ValueError(f"no readable image among {[f.name for f in files]}")
    return max(sizes, key=sizes.get)


def fetch_k8s(out_dir):
    out = out_dir / "k8s"
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        clone = Path(tmp) / "community"
        sparse_clone(K8S_REPO, "icons/png", clone)
        count = 0
        for folder, variant in K8S_SOURCES:
            base = clone / "icons" / "png" / folder / variant
            if not base.is_dir():
                continue
            groups = {}
            for f in base.glob("*.png"):
                groups.setdefault(f.stem.rsplit("-", 1)[0], []).append(f)
            for short, files in sorted(groups.items()):
                name = K8S_NAMES.get(short, short)
                shutil.copyfile(biggest(files), out / f"k8s-{name}.png")
                count += 1
    print(f"k8s: wrote {count} icons to {out}")


def find_browser():
    candidates = [shutil.which(n) for n in ("msedge", "chrome", "google-chrome", "chromium", "chromium-browser")]
    candidates += [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    ]
    return next((c for c in candidates if c and os.path.exists(c)), None)


def svg_to_png(browser, svg, png, size=512):
    with tempfile.TemporaryDirectory() as tmp:
        page = Path(tmp) / "p.html"
        shutil.copyfile(svg, Path(tmp) / "i.svg")
        page.write_text(
            f'<html><body style="margin:0;background:transparent"><img src="i.svg" '
            f'style="width:{size}px;height:{size}px;object-fit:contain;display:block"></body></html>')
        shot = Path(tmp) / "shot.png"
        run([browser, "--headless=new", "--disable-gpu", "--hide-scrollbars", "--default-background-color=00000000",
             f"--window-size={size},{size}", f"--screenshot={shot}", page.as_uri(),
             f"--user-data-dir={Path(tmp) / 'profile'}"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=90)
        if not shot.exists():
            raise RuntimeError("browser produced no screenshot")
        im = Image.open(shot).convert("RGBA")
        im.crop(im.getchannel("A").getbbox() or (0, 0) + im.size).save(png)


def fetch_docker(out_dir):
    browser = find_browser()
    if not browser:
        print("docker: needs Microsoft Edge or Chrome to convert the SVG logos; skipped", file=sys.stderr)
        return
    out = out_dir / "docker"
    out.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as tmp:
        clone = Path(tmp) / "devicon"
        sparse_clone(DEVICON_REPO, "icons", clone)
        count = 0
        for folder, svg, name in DOCKER_LOGOS:
            src = clone / "icons" / folder / svg
            if not src.exists():
                print(f"docker: {folder}/{svg} not found upstream, skipped")
                continue
            svg_to_png(browser, src, out / f"logo-{name}.png")
            count += 1
    print(f"docker: wrote {count} logos to {out}")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--only", choices=("k8s", "docker"))
    ap.add_argument("--dest", default=str(Path(__file__).resolve().parent.parent / "icons"))
    args = ap.parse_args()
    dest = Path(args.dest)
    if shutil.which("git") is None:
        print("git is required", file=sys.stderr)
        return 1
    if args.only in (None, "k8s"):
        fetch_k8s(dest)
    if args.only in (None, "docker"):
        fetch_docker(dest)
    return 0


if __name__ == "__main__":
    sys.exit(main())
