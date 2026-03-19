from src.docx_io.field_refresh import refresh_doc_fields_with_word


def test_refresh_write_back_replace_failure_keeps_original_file(tmp_path, monkeypatch):
    target_path = tmp_path / "output.docx"
    target_path.write_bytes(b"old-bytes")

    def _fake_refresh(shadow_path, _timeout_sec):
        shadow_path.write_bytes(b"new-bytes")
        return True, "ok(pywin32)"

    def _fail_replace(_src, _dst):
        raise PermissionError("replace blocked")

    monkeypatch.setattr("src.docx_io.field_refresh._refresh_via_pywin32", _fake_refresh)
    monkeypatch.setattr("src.docx_io.field_refresh.os.replace", _fail_replace)

    ok, detail = refresh_doc_fields_with_word(str(target_path), timeout_sec=10)

    assert ok is False
    assert "failed to write back" in detail
    assert target_path.read_bytes() == b"old-bytes"
