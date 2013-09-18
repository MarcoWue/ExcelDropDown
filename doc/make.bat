set name=ExcelDropDown

cd ..\src
doxygen doxyfile.doxy

cd ..\doc

if exist "html" (
	mkdir "chm"
	del /q "chm"
	copy "html\help.chm" "chm\%name%.chm"

	mkdir "htm"
	del /q "htm"

	del /q "html\class_*-members.html"
	copy "html\class_*.html" "htm\%name%.html"
	copy "html\doxygen.css" "htm"
	copy "html\nav_*.png" "htm"
	rmdir /s /q "html"
)

if exist "latex" (
	cd "latex"
	call "make.bat"
	cd ..

	mkdir "pdf"
	del /q "pdf"

	copy "latex\refman.pdf" "pdf\%name%.pdf"
	rmdir /s /q "latex"
)