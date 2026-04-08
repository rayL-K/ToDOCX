"""模板管理回归测试。"""

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


if __name__ == "__main__":
    unittest.main()
