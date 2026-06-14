import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


class LauncherRuntimeTests(unittest.TestCase):
    def test_launcher_uses_project_runtime_and_bootstraps_when_missing(self):
        launcher = (ROOT / "一键启动工具.bat").read_text(encoding="utf-8")

        self.assertIn(r".runtime\python\python.exe", launcher)
        self.assertIn("setup_runtime.ps1", launcher)
        self.assertIn("start_streamlit.ps1", launcher)
        self.assertNotIn("codex-runtimes", launcher)

    def test_bootstrap_installs_project_local_python_and_requirements(self):
        setup = (ROOT / "setup_runtime.ps1").read_text(encoding="utf-8")

        self.assertIn(".runtime", setup)
        self.assertIn("python.org", setup)
        self.assertIn("requirements.txt", setup)
        self.assertIn("-m pip install", setup)

    def test_streamlit_starter_uses_detached_shell_process(self):
        starter = (ROOT / "start_streamlit.ps1").read_text(encoding="utf-8")

        self.assertIn("ProcessStartInfo", starter)
        self.assertIn("UseShellExecute = $true", starter)
        self.assertIn("WindowStyle", starter)
        self.assertIn("Invoke-WebRequest", starter)
        self.assertIn("Streamlit did not become ready", starter)

    def test_runtime_dependencies_are_pinned(self):
        requirements = (ROOT / "requirements.txt").read_text(encoding="utf-8")

        self.assertIn("streamlit==1.58.0", requirements)
        self.assertIn("openpyxl==3.1.5", requirements)


if __name__ == "__main__":
    unittest.main()
