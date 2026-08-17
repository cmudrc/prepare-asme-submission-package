from __future__ import annotations

import json
import os
import shutil
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Inches
from PIL import Image, ImageDraw
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas


REPO = Path(__file__).resolve().parents[1]
SCRIPT = REPO / "scripts" / "prepare_submission.py"


def linework(path: Path, pixels: tuple[int, int], dpi: int = 900) -> None:
    image = Image.new("RGB", pixels, "white")
    draw = ImageDraw.Draw(image)
    margin = max(10, pixels[0] // 30)
    draw.rectangle((margin, margin, pixels[0] - margin, pixels[1] - margin), outline="black", width=max(3, pixels[0] // 400))
    draw.line((margin, pixels[1] - margin, pixels[0] // 2, margin, pixels[0] - margin, pixels[1] - margin), fill="black", width=max(3, pixels[0] // 450))
    draw.text((margin * 2, margin * 2), "DESIGN SPACE", fill="black")
    image.save(path, dpi=(dpi, dpi))


def word_fixture(path: Path, image_path: Path) -> None:
    document = Document()
    document.add_heading("A Validated Design Manuscript", 0)
    document.add_paragraph("A. Author, Carnegie Mellon University, author@example.edu")
    document.add_heading("Abstract", level=1)
    document.add_paragraph("This fixture validates final-production packaging.")
    document.add_paragraph("The feasible region is shown in Fig. 1 and remains legible at final size.")
    document.add_picture(str(image_path), width=Inches(3.0))
    document.add_paragraph("Fig. 1. A linework design-space diagram.")
    document.add_heading("Data Availability", level=1)
    document.add_paragraph("The fixture data are available at https://example.org/data.")
    document.add_heading("Conflict of Interest", level=1)
    document.add_paragraph("The authors declare no conflicts of interest.")
    document.add_heading("References", level=1)
    document.add_paragraph("[1] A. Author, 2026, A Fixture Reference.")
    document.save(path)


def complete_pdf(path: Path, image_path: Path) -> None:
    pdf = canvas.Canvas(str(path), pagesize=letter)
    pdf.setFont("Helvetica-Bold", 14)
    pdf.drawString(72, 740, "A Validated LaTeX Manuscript")
    pdf.setFont("Helvetica", 10)
    pdf.drawString(72, 718, "The feasible region is shown in Fig. 1 [1].")
    pdf.drawImage(str(image_path), 72, 480, width=216, height=108, preserveAspectRatio=True)
    pdf.drawString(72, 465, "Fig. 1. A linework design-space diagram.")
    pdf.drawString(72, 440, "Data Availability: https://example.org/data")
    pdf.drawString(72, 425, "Conflict of Interest: None.")
    pdf.drawString(72, 390, "References")
    pdf.drawString(72, 374, "[1] A. Author, A Fixture Reference, 2026.")
    pdf.save()


def latex_fixture(path: Path, image_path: Path, nocite: bool = False) -> None:
    main = r"""\documentclass{article}
\usepackage{graphicx}
\begin{document}
\title{A Validated LaTeX Manuscript}\maketitle
The feasible region in Fig.~\ref{fig:space} follows prior work~\cite{fixture}.
\begin{figure}
\centering
\includegraphics[width=3in]{figure1.png}
\caption{A linework design-space diagram.}
\label{fig:space}
\end{figure}
\section*{Data Availability}
The fixture data are available at \texttt{https://example.org/data}.
\section*{Conflict of Interest}
The authors declare no conflicts of interest.
%NOCITE%
\bibliographystyle{plain}
\bibliography{references}
\end{document}
"""
    main = main.replace("%NOCITE%", r"\nocite{*}" if nocite else "")
    bib = r"""@article{fixture,
  author = {Author, A.},
  title = {A Fixture Reference},
  year = {2026},
  journal = {Journal of Fixtures}
}
"""
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        (root / "main.tex").write_text(main, encoding="utf-8")
        (root / "references.bib").write_text(bib, encoding="utf-8")
        shutil.copy2(image_path, root / "figure1.png")
        with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as archive:
            for source in sorted(root.iterdir()):
                archive.write(source, source.name)


class PrepareSubmissionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp = tempfile.TemporaryDirectory(prefix="jmd-skill-test-")
        self.root = Path(self.temp.name)

    def tearDown(self) -> None:
        self.temp.cleanup()

    def run_packager(self, *arguments: str) -> subprocess.CompletedProcess[str]:
        environment = os.environ.copy()
        override = Path.home() / ".cache/codex-runtimes/codex-primary-runtime/dependencies/bin/override"
        environment["PATH"] = f"{override}:{environment.get('PATH', '')}"
        return subprocess.run(
            [sys.executable, str(SCRIPT), *arguments],
            text=True,
            capture_output=True,
            env=environment,
            check=False,
            timeout=90,
        )

    def test_word_package_preserves_text_and_extracts_figure(self) -> None:
        image = self.root / "figure.png"
        manuscript = self.root / "manuscript.docx"
        output = self.root / "word-output"
        linework(image, (2700, 1350))
        word_fixture(manuscript, image)

        result = self.run_packager(
            str(manuscript),
            "--paper-id", "MD-26-1001",
            "--journal", "jmd",
            "--output", str(output),
            "--require-section", "Data Availability",
            "--require-section", "Conflict of Interest",
        )

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertTrue((output / "MD-26-1001_submission_package.zip").is_file())
        self.assertTrue((output / "upload/MD-26-1001_complete.pdf").is_file())
        self.assertTrue((output / "upload/figures/Fig_1.tiff").is_file())
        text_only = output / "upload/MD-26-1001_text_only.docx"
        with zipfile.ZipFile(text_only) as archive:
            self.assertFalse(any(name.startswith("word/media/") for name in archive.namelist()))
            xml = archive.read("word/document.xml").decode("utf-8")
            self.assertIn("shown in Fig. 1", xml)
            self.assertIn("Fig. 1. A linework", xml)
        manifest = json.loads((output / "manifest.json").read_text())
        self.assertEqual(manifest["status"], "pass")
        self.assertGreaterEqual(manifest["figures"][0]["effective_dpi"], 899.9)
        self.assertTrue(list((output / "qa/complete").glob("*.png")))
        self.assertTrue(list((output / "qa/text-only").glob("*.png")))

    def test_low_effective_dpi_blocks_zip(self) -> None:
        image = self.root / "figure.png"
        manuscript = self.root / "manuscript.docx"
        output = self.root / "low-dpi-output"
        linework(image, (900, 450), dpi=300)
        word_fixture(manuscript, image)

        result = self.run_packager(
            str(manuscript),
            "--paper-id", "JCISE-26-1002",
            "--journal", "jcise",
            "--output", str(output),
        )

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertFalse((output / "JCISE-26-1002_submission_package.zip").exists())
        report = (output / "QA_REPORT.md").read_text()
        self.assertIn("LOW_EFFECTIVE_DPI", report)

    def test_overleaf_package_keeps_bib_separate_and_disables_graphics(self) -> None:
        image = self.root / "figure1.png"
        source = self.root / "overleaf.zip"
        pdf = self.root / "accepted.pdf"
        output = self.root / "latex-output"
        supplemental = self.root / "supplemental.txt"
        linework(image, (2700, 1350))
        latex_fixture(source, image)
        complete_pdf(pdf, image)
        supplemental.write_text("Synthetic supplemental material.\n", encoding="utf-8")

        result = self.run_packager(
            str(source),
            "--paper-id", "JCISE-26-1003",
            "--journal", "jcise",
            "--main", "main.tex",
            "--full-pdf", str(pdf),
            "--output", str(output),
            "--due-date", "2026-09-01",
            "--instruction", "Upload multipart figures separately",
            "--portal-rule", "Corresponding author uploads the complete set",
            "--supplemental-file", str(supplemental),
            "--require-section", "Data Availability",
        )

        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)
        self.assertTrue((output / "upload/native/references.bib").is_file())
        root_tex = (output / "upload/native/main.tex").read_text()
        self.assertIn("Production text-only override", root_tex)
        self.assertFalse((output / "upload/native/figure1.png").exists())
        self.assertTrue((output / "upload/figures/Fig_1.tiff").is_file())
        self.assertTrue((output / "upload/supplemental/supplemental.txt").is_file())
        self.assertTrue((output / "JCISE-26-1003_submission_package.zip").is_file())
        manifest = json.loads((output / "manifest.json").read_text())
        requirements = manifest["controlling_requirements"]
        self.assertEqual(requirements["due_date"], "2026-09-01")
        self.assertEqual(requirements["editorial_instructions"], ["Upload multipart figures separately"])
        self.assertEqual(requirements["supplemental_files"], ["supplemental.txt"])

    def test_nocite_star_blocks_package(self) -> None:
        image = self.root / "figure1.png"
        source = self.root / "overleaf.zip"
        pdf = self.root / "accepted.pdf"
        output = self.root / "nocite-output"
        linework(image, (2700, 1350))
        latex_fixture(source, image, nocite=True)
        complete_pdf(pdf, image)

        result = self.run_packager(
            str(source),
            "--paper-id", "JCISE-26-1004",
            "--main", "main.tex",
            "--full-pdf", str(pdf),
            "--output", str(output),
        )

        self.assertEqual(result.returncode, 2, result.stdout + result.stderr)
        self.assertIn("NOCITE_STAR", (output / "QA_REPORT.md").read_text())
        self.assertFalse((output / "JCISE-26-1004_submission_package.zip").exists())


if __name__ == "__main__":
    unittest.main()
