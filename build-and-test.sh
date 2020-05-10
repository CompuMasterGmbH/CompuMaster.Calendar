nuget restore CompuMaster.Calendar.VS2012.sln
nuget install NUnit.ConsoleRunner -Version 3.11.1
#msbuild /p:Configuration=Debug 
msbuild /p:Configuration=Debug /p:Platform="Any CPU" /p:PostBuildEvent="" CompuMaster.Calendar.VS2012.sln
#msbuild /p:Configuration=Release /p:Platform="Any CPU" /p:PostBuildEvent="" CompuMaster.Calendar.VS2012.sln
mono ./NUnit.ConsoleRunner.3.11.1/tools/nunit3-console.exe ./CompuMaster.Test.Calendar/bin/CompuMaster.Test.Calendar.dll
