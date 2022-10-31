 
run:
	dotnet run
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

test:
	/workspace/ConvertDocm/bin/Debug/net6.0/linux-x64/ConvertDocm