# gmcl\_speech

A module for Garry's Mod that provides speech recognition interfaces to developers.

## Compiling

This is Windows only (uses Microsoft's speech recognition capabilities).

The only supported compilation platform for this project is **Visual Studio 2017**. However, it's possible it'll work with *Visual Studio 2015* and *Visual Studio 2019* because of the unified runtime.

If stuff starts erroring or fails to work, be sure to check the correct line endings (`\n` and such) are present in the files.

## Requirements

This project requires [garrysmod\_common][1], a framework to facilitate the creation of compilations files (Visual Studio, make, XCode, etc). Simply set the environment variable `GARRYSMOD\_COMMON` or the premake option `--gmcommon=path` to the path of your local copy of [garrysmod\_common][1].

This project also needs [Windows 10 SDK](https://developer.microsoft.com/en-us/windows/downloads/windows-10-sdk/), which will provide the Windows and Speech API headers.

  [1]: https://github.com/danielga/garrysmod_common
