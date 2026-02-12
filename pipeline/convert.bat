@echo off
REM HWPX ↔ Markdown 변환 파이프라인 래퍼
REM 사용법:
REM   convert to-md  문서.hwpx [-o output.md]
REM   convert to-hwpx 문서.md [-o output.hwpx] [-r 원본.hwpx]
REM   convert smart  원본.hwpx 편집된.md [-o output.hwpx]

python "%~dp0convert.py" %*
