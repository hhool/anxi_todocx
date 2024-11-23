# todocx

A simple tool to convert xml and cvs files to docx files with a template docx file for anxi

## prerequisites

- vs2017 with c++ support
- [nuget](https://dist.nuget.org/win-x86-commandline/latest/nuget.exe) package manager for vs2017 or later
- c# 3.5 framework or higher

## how to use

1. build the project
2. run the todocx.exe with the following parameters
   - -s: the summary xml file
   - -i: the data list csv file
   - -t: the template docx file```

e.g.

todocx.exe -s summary.xml -i datalist.csv -t template.docx

## how to build

1. open the solution file todocx.sln
2. build the solution

## deploy with cmake

```bash
$ git clone url todocx.git
$ cd todocx
$ cmake . -B out\package -DCMAKE_BUILD_TYPE="Release"
$ cmake --build out\package --config Release
```

open explorer and go to `todocx\out\package` folder, you will see `todocx-x.x.x-win32.exe` file.

