if not os.is("windows") then
	error("Your operating system is not supported")
end

newoption({
	trigger = "gmcommon",
	description = "Sets the path to the garrysmod_common (https://bitbucket.org/danielga/garrysmod_common) directory",
	value = "path to garrysmod_common dir"
})

newoption({
	trigger = "speechsdk",
	description = "Sets the path to Microsoft's Speech Platform SDK x86 directory",
	value = "path to Speech SDK dir"
})

local gmcommon = _OPTIONS.gmcommon or os.getenv("GARRYSMOD_COMMON")
if gmcommon == nil then
	error("you didn't provide a path to your garrysmod_common (https://bitbucket.org/danielga/garrysmod_common) directory")
end

local speechsdk = _OPTIONS.speechsdk or os.getenv("SPEECH_SDK")
if speechsdk == nil then
	error("you didn't provide a path to Microsoft's Speech Platform SDK directory")
end

speechsdk = speechsdk:match("^(.+)[/\\]$") or speechsdk

include(gmcommon)

CreateSolution("speech")
	CreateProject(CLIENTSIDE)
		IncludeLuaShared()
		includedirs({speechsdk .. "/Include"})
		libdirs({speechsdk .. "/Lib"})
		links({"sapi"})