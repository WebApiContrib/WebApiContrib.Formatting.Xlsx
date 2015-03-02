Properties {
	$base_dir            = resolve-path .\..\..\
	$packages_dir        = "$base_dir\packages"
	$build_artifacts_dir = "$base_dir\build"
	$solution_name       = "$base_dir\WebApiContrib.Formatting.Xlsx.sln"
	$test_dll            = "$build_artifacts_dir\WebApiContrib.Formatting.Xlsx.Tests.dll"
	$nuget_exe           = "$base_dir\.nuget\Nuget.exe"
	$test_result_path    = "$base_dir\TestResults\"
	$test_result_file    = [System.IO.Path]::Combine($test_result_path, "TestResults.trx")
	$vscomntools_path    = VSCommonToolsPath
	$mstest_path         = (Get-Item $vscomntools_path).Parent.FullName
	$mstest_exe          = [System.IO.Path]::Combine($mstest_path, "IDE\MSTest.exe")
}

Task Default -Depends BuildWebApiContrib, RunUnitTests, NuGetBuild

Task BuildWebApiContrib -Depends Clean, Build

Task Clean {
	Exec { msbuild $solution_name /v:Quiet /t:Clean /p:Configuration=Release }
}

Task Build -depends Clean {
	Exec { msbuild $solution_name /v:Quiet /t:Build /p:Configuration=Release /p:OutDir=$build_artifacts_dir\ } 
}

Task NuGetBuild -depends Clean {
	& $nuget_exe pack "$base_dir/src/WebApiContrib.Formatting.Xlsx/WebApiContrib.Formatting.Xlsx.csproj" -Build -OutputDirectory $build_artifacts_dir -Verbose -Properties Configuration=Release
}

Task RunUnitTests -depends Build {
    New-Item -ItemType Directory -Force -Path "$test_result_path"
    $test_arguments = @("/resultsFile:$test_result_file")
    $test_arguments += "/testcontainer:$test_dll"
	
	If (Test-Path $test_result_file) { Remove-Item $test_result_file }

    # psake will terminate the execution if mstest throws an exception.
    Exec { & $mstest_exe $test_arguments }
}

# Find VS Common Tools directory path for the most recent version of Visual Studio.
Function VSCommonToolsPath {
	If (Test-Path Env:VS150COMNTOOLS) { Return (Get-ChildItem env:VS150COMNTOOLS).Value }
	If (Test-Path Env:VS140COMNTOOLS) { Return (Get-ChildItem env:VS140COMNTOOLS).Value }
	If (Test-Path Env:VS130COMNTOOLS) { Return (Get-ChildItem env:VS130COMNTOOLS).Value }
	If (Test-Path Env:VS120COMNTOOLS) { Return (Get-ChildItem env:VS120COMNTOOLS).Value }
	If (Test-Path Env:VS110COMNTOOLS) { Return (Get-ChildItem env:VS110COMNTOOLS).Value }
	If (Test-Path Env:VS100COMNTOOLS) { Return (Get-ChildItem env:VS100COMNTOOLS).Value }
	If (Test-Path Env:VS90COMNTOOLS ) { Return (Get-ChildItem env:VS90COMNTOOLS).Value  }
	If (Test-Path Env:VS80COMNTOOLS ) { Return (Get-ChildItem env:VS80COMNTOOLS).Value  }
}