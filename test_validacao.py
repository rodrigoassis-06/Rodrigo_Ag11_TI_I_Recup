import unittest
import os
import docx
from PIL import Image

class TestEntregaModelagem(unittest.TestCase):
    def setUp(self):
        self.base_dir = os.path.dirname(os.path.abspath(__file__))

    def test_arquivos_obrigatorios_existem(self):
        arquivos = [
            "Rodrigo_Ag11_TI_I_Recup.docx",
            "diagrama_caso_de_uso.png",
            "interacao_ia.png",
            "diagrama.html",
            "gerar_relatorio.py"
        ]
        for arq in arquivos:
            caminho = os.path.join(self.base_dir, arq)
            self.assertTrue(os.path.exists(caminho), f"Arquivo ausente: {arq}")
            self.assertGreater(os.path.getsize(caminho), 0, f"Arquivo vazio: {arq}")

    def test_dimensoes_imagens(self):
        diag_path = os.path.join(self.base_dir, "diagrama_caso_de_uso.png")
        ia_path = os.path.join(self.base_dir, "interacao_ia.png")

        with Image.open(diag_path) as img:
            self.assertGreater(img.width, 800)
            self.assertGreater(img.height, 600)

        with Image.open(ia_path) as img:
            self.assertGreater(img.width, 600)
            self.assertGreater(img.height, 400)

    def test_conteudo_docx(self):
        docx_path = os.path.join(self.base_dir, "Rodrigo_Ag11_TI_I_Recup.docx")
        doc = docx.Document(docx_path)
        self.assertGreater(len(doc.paragraphs), 10)
        self.assertGreater(len(doc.tables), 0)

if __name__ == "__main__":
    unittest.main()
