from core import config


def test_id_helpers_preserve_three_digits() -> None:
    assert config.dir_name("001") == "officeval_001"
    assert config.verifier_filename("001") == "officeval_001_verifier.py"
    assert config.id_from_dir_name("officeval_001") == "001"
    assert config.id_from_verifier_filename("officeval_001_verifier.py") == "001"
    assert config.id_from_dir_name("officeval_1") is None



def test_allowed_document_extensions_exclude_legacy_office_formats() -> None:
    assert config.ALLOWED_DOC_EXTENSIONS == frozenset({
        ".docx", ".xlsx", ".xlsm", ".pptx", ".pdf",
    })
    assert not ({".doc", ".ppt", ".xls"} & config.ALLOWED_DOC_EXTENSIONS)


def test_com_policy_is_explicit_and_disjoint() -> None:

    config.validate_config()
    assert config.COM_REQUIRED_VERIFIER_IDS == frozenset({
        "011", "023", "039", "081",
    })
    assert config.COM_VERIFIER_IDS == config.COM_REQUIRED_VERIFIER_IDS
    assert not (
        config.COM_REQUIRED_VERIFIER_IDS
        & config.COM_FORCED_NORMAL_VERIFIER_IDS
    )
