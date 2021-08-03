
dotnet pack -p:PackageVersion=$args
dotnet nuget push .\bin\Debug\NPOIHelper.$args.nupkg -k oy2gveeqrojs7cfjk2uqzdm53d4a4juevcbspphmdjy3c4 -s https://www.nuget.org
