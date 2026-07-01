from . import austria, czech, hungary, korea, netherlands, poland, slovakia

DEFAULT_IMPORT_FORMAT = "한국 (Korea)"
IMPORT_FORMATS = {
    "한국 (Korea)": korea.build_records,
    "네덜란드 (Netherlands)": netherlands.build_records,
    "오스트리아 (Austria)": austria.build_records,
    "헝가리 (Hungary)": hungary.build_records,
    "폴란드 (Poland)": poland.build_records,
    "체코 (Czech Republic)": czech.build_records,
    "슬로바키아 (Slovakia)": slovakia.build_records,
}


def build_source_records(sheet, import_format, progress_callback=None):
    try:
        parser = IMPORT_FORMATS[import_format]
    except KeyError as exc:
        raise ValueError(f"지원하지 않는 국가 양식입니다: {import_format}") from exc
    return parser(sheet, progress_callback)
