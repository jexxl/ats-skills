# /// script
# requires-python = ">=3.11"
# dependencies = [
#   "pillow>=11.2.1",
#   "playwright>=1.52.0",
# ]
# ///
"""Render an ATS algorithm cover HTML file to a transparent high-DPI PNG."""

from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import Page, sync_playwright


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Render an ATS algorithm cover card to PNG.")
    parser.add_argument("html", type=Path, help="Path to cover.html.")
    parser.add_argument("output", type=Path, help="Path to output PNG.")
    parser.add_argument("--viewport-width", type=int, default=600)
    parser.add_argument("--viewport-height", type=int, default=300)
    parser.add_argument("--scale", type=float, default=3.0, help="Device scale factor.")
    parser.add_argument("--dpi", type=int, default=300, help="PNG DPI metadata.")
    parser.add_argument(
        "--browser-channel",
        default="chrome",
        help="Playwright browser channel, such as chrome or msedge. Use an empty string for bundled Chromium.",
    )
    parser.add_argument(
        "--opaque-background",
        action="store_true",
        help="Keep the HTML page background instead of making rounded corners transparent.",
    )
    parser.add_argument("--timeout-ms", type=int, default=30_000)
    return parser.parse_args()


def wait_for_mermaid(page: Page, timeout_ms: int) -> None:
    page.wait_for_function("() => Boolean(window.mermaid)", timeout=timeout_ms)
    page.wait_for_function(
        """() => {
            const card = document.querySelector(".algorithm-card");
            const svg = document.querySelector(".mermaid svg");
            return Boolean(card && svg && svg.getBBox().width > 0 && svg.getBBox().height > 0);
        }""",
        timeout=timeout_ms,
    )


def set_png_dpi(path: Path, dpi: int) -> None:
    with Image.open(path) as image:
        image.save(path, dpi=(dpi, dpi))


def make_page_background_transparent(page: Page) -> None:
    page.add_style_tag(
        content="""
        html,
        body {
            background: transparent !important;
        }
        """
    )


def launch_chromium(playwright, channel: str):
    launch_args = ["--no-sandbox", "--disable-dev-shm-usage"]
    if channel:
        try:
            return playwright.chromium.launch(
                channel=channel,
                headless=True,
                args=launch_args,
            )
        except PlaywrightError:
            pass

    return playwright.chromium.launch(headless=True, args=launch_args)


def render_cover(args: argparse.Namespace) -> None:
    html_path = args.html.resolve()
    output_path = args.output.resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = launch_chromium(p, args.browser_channel)
        page = browser.new_page(
            viewport={"width": args.viewport_width, "height": args.viewport_height},
            device_scale_factor=args.scale,
        )
        page.goto(html_path.as_uri(), wait_until="domcontentloaded", timeout=args.timeout_ms)
        wait_for_mermaid(page, args.timeout_ms)
        if not args.opaque_background:
            make_page_background_transparent(page)
        page.locator(".algorithm-card").screenshot(
            path=str(output_path),
            omit_background=not args.opaque_background,
        )
        browser.close()

    set_png_dpi(output_path, args.dpi)


def main() -> None:
    render_cover(parse_args())


if __name__ == "__main__":
    main()
