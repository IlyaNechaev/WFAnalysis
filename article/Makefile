FILE_MD = .\\report.md
FILE_DOCX = .\\build\\report.docx
FILE_REFERENCE = .\\custom-reference.docx
FILE_TEMPLATE = .\\template.docx

all: docx

docx: $(FILE_DOCX)

$(FILE_DOCX): $(FILE_MD) $(FILE_REFERENCE)
	powershell -command ./build.ps1 \
	-md $(FILE_MD) \
	-reference $(FILE_REFERENCE) \
	-docx $(FILE_DOCX) \
	-template $(FILE_TEMPLATE)

clean:
	del $(FILE_DOCX)