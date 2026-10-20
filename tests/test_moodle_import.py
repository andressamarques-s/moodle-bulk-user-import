import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import moodle_import as mi

COURSES = ["Onboarding", "First Aid"]


def row(**overrides):
    base = {"org": "acme", "unit": "Head Office", "name": "Maria Silva",
            "cpf": "529.982.247-25", "email": "maria@example.com", "course": "Onboarding"}
    return {**base, **overrides}


def test_valid_cpf():
    assert mi.is_valid_cpf("52998224725")
    assert not mi.is_valid_cpf("52998224724")
    assert not mi.is_valid_cpf("11111111111")


def test_normalize_cpf_keeps_leading_zeros():
    assert mi.normalize_cpf(1234567890.0) == "01234567890"
    assert mi.normalize_cpf("529.982.247-25") == "52998224725"


def test_valid_row_has_no_errors():
    (p,) = mi.build_people([row()], COURSES)
    assert p.errors == []
    assert (p.org, p.firstname, p.lastname) == ("ACME", "Maria", "Silva")


def test_validation_errors():
    (p,) = mi.build_people([row(cpf="123", email="bad", course="Nope", name="Maria")], COURSES)
    assert "Invalid CPF" in p.errors
    assert "Invalid e-mail" in p.errors
    assert "Course is not in the Moodle course list" in p.errors
    assert "Enter first and last name" in p.errors


def test_duplicated_cpf_and_mixed_orgs():
    people = mi.build_people([row(), row(org="other", name="Joao Souza")], COURSES)
    assert any("Duplicated CPF" in e for e in people[1].errors)
    assert any("same organization" in e for e in people[1].errors)


def test_empty_rows_are_skipped():
    assert mi.build_people([{"org": "ACME"}], COURSES) == []


def test_csv_output():
    people = mi.build_people([row()], COURSES)
    lines = mi.generate_csv(people, "pw", ";").decode().splitlines()
    assert lines[0] == "username;firstname;lastname;email;password;institution;course1;group1"
    assert lines[1] == "52998224725;Maria;Silva;maria@example.com;pw;ACME/HEAD OFFICE;Onboarding;HEAD OFFICE"


def test_save_files_does_not_overwrite(tmp_path):
    folders = {k: tmp_path / k for k in ("pdf", "csv", "xlsx")}
    people = mi.build_people([row()], COURSES)
    first = mi.save_files(people, folders, "pw")
    second = mi.save_files(people, folders, "pw")
    assert first["csv"].name != second["csv"].name
    assert "(2)" in second["csv"].name
    assert all(p.stat().st_size > 0 for p in {**first, **second}.values())
