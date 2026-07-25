from core.verifier_registry import build_registry


def test_registry_contains_all_verifiers() -> None:
    registry = build_registry()
    assert len(registry) == 100
    assert registry.ids()[0] == "001"
    assert registry.ids()[-1] == "100"
    assert all(registry.path_for(id_).is_file() for id_ in registry.ids())
