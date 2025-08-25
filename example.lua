require("speech")

local SPEECH = speech.create("HKEY_LOCAL_MACHINE\\SOFTWARE\\Microsoft\\Speech\\Recognizers\\Tokens\\MS-1033-80-DESK")
-- 38 stands for SPEI_RECOGNITION
print("speech interest", SPEECH:interest(38))
print("speech grammar", SPEECH:grammar("example_grammar.xml"))
print("speech grammar state", SPEECH:grammar_state("enabled"))
print("speech rule state", SPEECH:rule_state(nil, "active"))
print("speech resume", SPEECH:resume())

local spellnames = {
	["die"] = "fireball",
}
for spellname, spell in pairs(Arcane.RegisteredSpells) do
	spellnames[string.gsub(spellname, "_", " ")] = spell.id
end

hook.Add("Think", "Arcane Stuff", function()
	local num, events, err = SPEECH:events(10)
	if num == nil then
		print("ERROR", err)
		return
	end

	if num == 0 then
		return
	end

	for _, event in pairs(events) do
		local text = event.text
		if text ~= nil then
			for spellname, spellid in pairs(spellnames) do
				if string.find(text, string.lower(spellname)) then
					RunConsoleCommand("arcana", spellid)
				end
			end
		end
	end
end)
