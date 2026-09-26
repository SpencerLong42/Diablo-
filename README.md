# Diablo 4 build notes

## Season 15 – Warlock "Blazing Scream": level 70 → endgame playbook

- **[Season-15-Warlock-Blazing-Scream-Guide.docx](Season-15-Warlock-Blazing-Scream-Guide.docx)**: the formatted Word version, about 30 pages with a title page, clickable contents and tables.
- **[Season-15-Warlock-Blazing-Scream-Guide.md](Season-15-Warlock-Blazing-Scream-Guide.md)**: the same content for reading on GitHub.

The guide starts from **level 70 with Rare items only and the Splinter of Destruction**. It covers:

1. **The short answer to "should I just run Lair Bosses?"**, a 5-phase overview and a master checklist.
2. **A step-by-step endgame plan** from Rares to Mythics: what to do in town, then at Penitent → T1 → T5 → T8 → T12.
3. **Stat priority for every gear slot** before you have the ideal items. The lists come from the Season 15 game data (what can actually roll on each Warlock slot). Each slot also lists its aspect, tempering manual, sockets and what to ignore. There's also an aspect slot-rules table, how to fix Rares (Codex, Occultist, Horadric Cube), and what gems, Soul Splinters, runes and charms to use in the meantime.
4. **Season 15 systems:** the questline, all 24 Splinter of Destruction ranks, the Season Rank ladder, War Plans, Talismans, mercenaries, and the Pit/Torment tiers.
5. **The build in phases:**
   - Starter (no uniques)
   - Midgame (Elegy + Infernal Homunculus + Temerity)
   - Endgame, Push and Speedfarm
   - The Orbital Scream / Teleport version

   Each phase lists skill points, gear per slot, tempers, masterwork targets and sockets. Paragon boards and glyphs are covered too.
6. **Where every item drops**, crafting (masterwork, runewords, Mythics), survival notes and the repeating farm loop.

Build reference: <https://mobalytics.gg/diablo-4/builds/warlock-blazing-scream>. The detailed setups come from Maxroll's decoded planner and D4Guides' version of the same build; see the guide's notes and sources.

### Regenerating the guide
Both files are generated from `guide-source/content1-4.js`:

```bash
cd guide-source && npm install && node build.js   # writes the .docx and .md into the repo root
```
