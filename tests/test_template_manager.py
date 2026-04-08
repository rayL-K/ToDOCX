"""模板管理回归测试。"""

import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest

from src.template_manager import TemplateManager


class TemplateManagerTests(unittest.TestCase):
    def test_rename_template_preserves_description(self):
        with TemporaryDirectory() as tmp:
            manager = TemplateManager(tmp)
            manager.save_template("原模板", {"body": {"font_name_cn": "宋体"}}, description="旧说明")

            renamed = manager.rename_template("原模板", "新模板")

            self.assertTrue(renamed)
            templates = {item["name"]: item for item in manager.list_templates()}
            self.assertIn("新模板", templates)
            self.assertEqual(templates["新模板"]["description"], "旧说明")

    def test_save_template_rejects_empty_safe_name(self):
        with TemporaryDirectory() as tmp:
            manager = TemplateManager(tmp)

            with self.assertRaises(ValueError):
                manager.save_template("///", {"body": {}})

    def test_list_templates_marks_corrupted_file(self):
        with TemporaryDirectory() as tmp:
            template_dir = Path(tmp)
            manager = TemplateManager(tmp)
            (template_dir / "broken.json").write_text("{invalid", encoding="utf-8")

            templates = manager.list_templates()

            broken = next(item for item in templates if item["name"] == "broken")
            self.assertEqual(broken["status"], "corrupted")
            self.assertIn("损坏", broken["description"])

    def test_list_templates_marks_structurally_invalid_template(self):
        with TemporaryDirectory() as tmp:
            template_dir = Path(tmp)
            manager = TemplateManager(tmp)
            (template_dir / "bad-shape.json").write_text(
                json.dumps({"name": "坏模板", "styles": []}, ensure_ascii=False),
                encoding="utf-8",
            )

            templates = manager.list_templates()

            broken = next(item for item in templates if item["name"] == "bad-shape")
            self.assertEqual(broken["status"], "corrupted")

    def test_save_template_writes_schema_version(self):
        with TemporaryDirectory() as tmp:
            template_dir = Path(tmp)
            manager = TemplateManager(tmp)

            manager.save_template("学术模板", {"body": {"font_name_cn": "宋体"}}, description="示例")

            data = json.loads((template_dir / "学术模板.json").read_text(encoding="utf-8"))
            self.assertEqual(data["schema_version"], 1)
            self.assertEqual(data["description"], "示例")

    def test_migrate_legacy_templates_copies_missing_files(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            legacy_dir = root / "legacy"
            new_dir = root / "new"
            legacy_dir.mkdir()
            manager = TemplateManager(str(new_dir))
            manager.legacy_template_dir = legacy_dir

            (legacy_dir / "旧模板.json").write_text(
                json.dumps({"name": "旧模板", "styles": {}, "description": ""}, ensure_ascii=False),
                encoding="utf-8",
            )

            manager._migrate_legacy_templates()

            self.assertTrue((new_dir / "旧模板.json").exists())

    def test_migrate_legacy_templates_skips_invalid_template(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            legacy_dir = root / "legacy"
            new_dir = root / "new"
            legacy_dir.mkdir()
            manager = TemplateManager(str(new_dir))
            manager.legacy_template_dir = legacy_dir

            (legacy_dir / "坏模板.json").write_text(
                json.dumps({"name": "坏模板", "styles": []}, ensure_ascii=False),
                encoding="utf-8",
            )

            manager._migrate_legacy_templates()

            self.assertFalse((new_dir / "坏模板.json").exists())


if __name__ == "__main__":
    unittest.main()
