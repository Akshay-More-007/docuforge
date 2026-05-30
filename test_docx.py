from dotenv import load_dotenv
load_dotenv()

from src.document.extractor import extract_docx
from src.document.style_parser import extract_styles

content = extract_docx("C:/Users/Batman/Downloads/Prep/Part10_Final_CompanyPrep_CheatSheet.docx")
styles  = extract_styles("C:/Users/Batman/Downloads/Prep/Part10_Final_CompanyPrep_CheatSheet.docx")

print(content["sections"][:2])
print(styles["heading_styles"])