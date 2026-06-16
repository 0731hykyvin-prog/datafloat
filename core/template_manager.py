import os
import json
import shutil
import sys

APP_NAME = "Datafloat"


def _resource_dir():
    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return os.path.join(sys._MEIPASS, "templates")
    return os.path.join(os.getcwd(), "templates")


def _writable_template_dir():
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        exe_templates = os.path.join(exe_dir, "templates")
        try:
            os.makedirs(exe_templates, exist_ok=True)
            test_file = os.path.join(exe_templates, ".write_test")
            with open(test_file, "w", encoding="utf-8") as f:
                f.write("ok")
            os.remove(test_file)
            return exe_templates
        except OSError:
            appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
            return os.path.join(appdata, APP_NAME, "templates")

    return os.path.join(os.getcwd(), "templates")


class TemplateManager:

    def __init__(self):
        self.bundled_dir = _resource_dir()
        self.template_dir = _writable_template_dir()

        os.makedirs(self.template_dir, exist_ok=True)
        self._copy_default_templates()

    def _copy_default_templates(self):
        if not os.path.isdir(self.bundled_dir):
            return

        for file_name in os.listdir(self.bundled_dir):
            if not file_name.endswith(".json"):
                continue

            source = os.path.join(self.bundled_dir, file_name)
            target = os.path.join(self.template_dir, file_name)
            if not os.path.exists(target):
                shutil.copy2(source, target)

    # 保存模板
    def save_template(self, name, mapping):

        path = os.path.join(
            self.template_dir,
            f"{name}.json"
        )

        with open(path, "w", encoding="utf-8") as f:

            json.dump(
                mapping,
                f,
                ensure_ascii=False,
                indent=4
            )

    # 读取模板
    def load_template(self, name):

        path = os.path.join(
            self.template_dir,
            f"{name}.json"
        )

        if not os.path.exists(path):
            return None

        with open(path, "r", encoding="utf-8") as f:

            return json.load(f)

    # 获取所有模板
    def list_templates(self):

        if not os.path.exists(self.template_dir):
            return []

        return sorted([
            f.replace(".json", "")
            for f in os.listdir(self.template_dir)
            if f.endswith(".json")
        ])
