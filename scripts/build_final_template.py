# -*- coding: utf-8 -*-
"""ساخت دوباره‌پذیر فایل اکسل نهایی حافا در دو مرحله.

اجرا از ریشه مخزن:
    python scripts/build_final_template.py
"""
from fill_template_from_extraction import main as fill_from_primary_sources
from finalize_template_with_benchmarks import finalize


def main() -> None:
    fill_from_primary_sources()
    finalize()


if __name__ == "__main__":
    main()
