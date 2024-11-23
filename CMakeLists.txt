# CMakeLists.txt for todocx project
# Author: hhool
# Date: 2024-11-23

# Set the minimum required version of CMake
cmake_minimum_required(VERSION 3.0)

# Set the project name
project(todocx VERSION 1.0.0 LANGUAGES CSharp)

# Include CMake utilities for CSharp, for WinForm and WPF application support.
include(CSharpUtilities)

# Configure the build type
if(NOT CMAKE_BUILD_TYPE)
    set(CMAKE_BUILD_TYPE Debug)
endif()

# set output directory
if(CMAKE_SOURCE_DIR STREQUAL PROJECT_SOURCE_DIR)
    set(CMAKE_RUNTIME_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/bin)
    set(CMAKE_ARCHIVE_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
    set(CMAKE_LIBRARY_OUTPUT_DIRECTORY ${CMAKE_BINARY_DIR}/lib)
endif()

# Set the build type
set(CMAKE_CONFIGURATION_TYPES "Debug;Release")

# Set the .NET Framework version
set(CMAKE_DOTNET_TARGET_FRAMEWORK_VERSION "v3.5.0")

# Set the source files
set(SOURCES
    ${CMAKE_CURRENT_SOURCE_DIR}/todocx/App.config
    ${CMAKE_CURRENT_SOURCE_DIR}/todocx/Program.cs
    ${CMAKE_CURRENT_SOURCE_DIR}/todocx/packages.config
    ${CMAKE_CURRENT_SOURCE_DIR}/todocx/csv2docx.cs
    ${CMAKE_CURRENT_SOURCE_DIR}/todocx/Properties/AssemblyInfo.cs
)

# Add the executable target
add_executable(todocx ${SOURCES})

# set arch for the executable target
set_property(TARGET todocx PROPERTY WIN32_EXECUTABLE TRUE)
target_compile_options(todocx PRIVATE "/platform:x86" )

# Set the executable target properties console
set_property(TARGET todocx PROPERTY VS_DOTNET_TARGET_FRAMEWORK_VERSION "v3.5")

# Set the output directory for the executable target
set_property(TARGET todocx PROPERTY DOTNET_OUTPUT_DIRECTORY ${CMAKE_DOTNET_OUTPUT_DIRECTORY})

# Set the startup project
set_property(DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR} PROPERTY VS_STARTUP_PROJECT todocx)

# Set the project properties
set_property(TARGET todocx PROPERTY CONSOLE_APPLICATION TRUE)

# Set the project references
set_property(TARGET todocx PROPERTY VS_DOTNET_REFERENCES
                "System"
                "System.Core"
                "System.Data"
                "System.Data.DataSetExtensions"
                "System.Xml"
                "System.Xml.Linq"
                "WindowsBase")
# Find Nuget (install the latest CLI here: https://www.nuget.org/downloads).
find_program(NUGET nuget)
if(NOT NUGET)
    message(FATAL "CMake could not find the nuget command line tool. Please install it!")
else()
    # Copy the Nuget config file from source location to the CMake build directory.
    configure_file(${CMAKE_CURRENT_SOURCE_DIR}\\todocx\\packages.config.in ${CMAKE_CURRENT_SOURCE_DIR}\\todocx\\packages.config COPYONLY)
    # Run Nuget using the .config file to installing any missing dependencies to the build directory.
    execute_process(COMMAND 
        ${NUGET} restore ${CMAKE_CURRENT_SOURCE_DIR}\\todocx\\packages.config -SolutionDirectory ${CMAKE_BINARY_DIR}
        WORKING_DIRECTORY ${CMAKE_BINARY_DIR}
    )
endif()


# Set VS Package references
set_property(TARGET todocx PROPERTY VS_DOTNET_REFERENCE_DocumentFormat.OpenXml
                ${CMAKE_BINARY_DIR}/packages/Open-XML-SDK.2.9.1/lib/net35/DocumentFormat.OpenXml.dll)

# Set the install directory
if(NOT CMAKE_INSTALL_PREFIX)
    set(CMAKE_INSTALL_PREFIX "C:/Program Files (x86)/todocx")
endif()

# Install the executable target
install(TARGETS todocx
    DESTINATION ${CMAKE_INSTALL_PREFIX}
)

# Install the project files
install(FILES
    ${CMAKE_BINARY_DIR}/todocx.exe
    ${CMAKE_BINARY_DIR}/todocx.pdb
    DESTINATION ${CMAKE_INSTALL_PREFIX}
)


