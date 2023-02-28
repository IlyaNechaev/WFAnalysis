FILE_MD = .\\report.md
FILE_DRAFT_MD = .\\draft.md
FILE_DOCX = .\\build\\report.docx
FILE_REFERENCE = .\\custom-reference.docx
FILE_TEMPLATE = .\\template.docx
FILE_BIBTEX = .\\articles.bib

all: docx

docx: $(FILE_DOCX)

$(FILE_DOCX): $(FILE_MD) $(FILE_REFERENCE) $(FILE_BIBTEX)
	powershell -command ./build.ps1 \
	-md $(FILE_MD) \
	-reference $(FILE_REFERENCE) \
	-docx $(FILE_DOCX) \
	-template $(FILE_TEMPLATE)

$(FILE_MD): $(FILE_DRAFT_MD)
	powershell -Command "(Get-Content -Raw $(FILE_DRAFT_MD)) -Replace '`draft_start`((?s).*?)`draft_end`',$$1 -Replace '!>.*\n','' | Out-File $(FILE_MD) -Encoding default"

clean:
	del $(FILE_DOCX)