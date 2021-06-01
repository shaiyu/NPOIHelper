
dotnet pack -p:PackageVersion=$args
dotnet nuget push .\bin\Debug\NPOIHelper.$args.nupkg -k key -s https://www.nuget.org
