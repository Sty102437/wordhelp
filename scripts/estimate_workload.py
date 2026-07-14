#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Estimate workload and select engine — cross-platform.

Scoring:
  paragraphs: min(P/10, 5)
  cells:      min(C/5, 5)
  images:     min(I/3, 5)
  TOC:        +2
  academic/government: +3

  score > 8  -> minimax-docx (heavy engine)
  score <= 8 -> python-docx (light engine)

Usage: python estimate_workload.py -P 50 -C 30 [-T 5] [-I 0] [--type academic] [--toc]
"""
import argparse


def estimate(paragraphs, cells, tables=0, images=0, doc_type="general", toc=False):
    score = 0.0
    score += min(paragraphs / 10, 5)
    score += min(cells / 5, 5)
    score += min(images / 3, 5)
    if toc:
        score += 2
    if doc_type in ("academic", "government"):
        score += 3

    if score > 8:
        engine = "minimax-docx"
    else:
        engine = "python-docx"

    return engine, score


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Estimate workload and select engine")
    parser.add_argument("-P", type=int, required=True, help="Paragraph count")
    parser.add_argument("-C", type=int, required=True, help="Total table cell count")
    parser.add_argument("-T", type=int, default=0, help="Table count (default: 0)")
    parser.add_argument("-I", type=int, default=0, help="Image count (default: 0)")
    parser.add_argument("--type", default="general", choices=["general", "academic", "government"],
                        help="Document type (default: general)")
    parser.add_argument("--toc", action="store_true", help="TOC needed")
    args = parser.parse_args()

    engine, score = estimate(args.P, args.C, args.T, args.I, args.type, args.toc)
    print(f"engine={engine}; score={score}")
