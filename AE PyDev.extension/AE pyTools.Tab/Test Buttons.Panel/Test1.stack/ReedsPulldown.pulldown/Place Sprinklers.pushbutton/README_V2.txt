Basically the first test I ran was garbage so here's test 2

Make rules JSON files for identifying rooms, then what each discipline should do in that type of room, I started with lighting.

Next in lib we have the rules_loader that reads our rules, classifies each room then applies the rules by discipline.

Finally the actual script imports our rules_loader, collects the rooms with various methods that could probably be in another script (5)
proposes a grid to place things (6) and then (7) selects our lighting fixture that should probably also be in another script under TOOLS with a name like
select_lighting_fixtures. (8) is a small helper that could go in another very small script titled "isHosted" (9) actually places our fixtures per the grid that was created in (6). and (10) our main method is just printing helpful statements letting us know what happened.
