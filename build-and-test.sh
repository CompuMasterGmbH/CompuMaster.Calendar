nuget restore CompuMaster.Calendar.VS2012.sln
nuget install NUnit.ConsoleRunner -Version 3.12.0
#msbuild /p:Configuration=Debug 
msbuild /p:Configuration=Debug /p:Platform="Any CPU" /p:PostBuildEvent="" CompuMaster.Calendar.VS2012.sln
#msbuild /p:Configuration=Release /p:Platform="Any CPU" /p:PostBuildEvent="" CompuMaster.Calendar.VS2012.sln
curl https://raw.githubusercontent.com/jochenwezel/nunit-transforms-to-junit-gitlab-compatible/master/nunit3-junit/nunit3-junit.xslt > nunit3-junit.xslt
mono ./NUnit.ConsoleRunner.3.12.0/tools/nunit3-console.exe ./CompuMaster.Test.Calendar/bin/CompuMaster.Test.Calendar.dll --result=TestResult.xml "--result=TestResult.JUnit.xml;transform=nunit3-junit.xslt"
rm nunit3-junit.xslt
