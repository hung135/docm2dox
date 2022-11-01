 
run:
	#dotnet run ./example.docm ./output.docx
	rm *.docm -f
	cp ../ChangeBeforePrinting.docm ./
	dotnet run ./ChangeBeforePrinting.docm ./output.docx
win10: clean
	dotnet build -r win10-x64
add_pkg:
	dotnet add package DocumentFormat.OpenXml --version 2.5.0
clean:
	rm -rf bin
	rm -rf obj

publish: win10
	dotnet publish -r linux-x64 --self-contained true
	dotnet publish -r win10-x64 --self-contained true

publish_debug: win10
	dotnet publish -r linux-x64 --self-contained false
	dotnet publish -r win10-x64 --self-contained false
test:
	/workspace/ConvertDocm/bin/Debug/net6.0/linux-x64/ConvertDocm ./test.doc output.docx
get_sample:
	wget https://example-files.online-convert.com/document/docm/example_no_vba.docm
	wget https://www.activebarcode.com/download/demo/ChangeBeforePrinting.docm