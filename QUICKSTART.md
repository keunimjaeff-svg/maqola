# Quick Start Guide

## TL;DR

```bash
npm install
node build_final_local.js
```

Output: `./outputs/Saliev_Thermoelasticity_IMRAD.docx`

---

## What This Does

Creates a scientific article in Word format about **thermoelasticity** with:
- ✅ IMRAD structure (Introduction, Methods, Results, Discussion)
- ✅ 8 complex mathematical equations in native Word format
- ✅ Russian language content
- ✅ Proper academic formatting
- ✅ References and citations

## Files Overview

| File | Purpose |
|------|---------|
| `build_final_local.js` | **Use this** - outputs to `./outputs/` directory |
| `build_final.js` | Original version - outputs to `/mnt/user-data/outputs/` |
| `package.json` | Node.js dependencies configuration |
| `setup_and_build.sh` | Automated setup script (Linux/Mac) |
| `README.md` | Detailed documentation |
| `USAGE_INSTRUCTIONS.md` | Complete usage guide |

## Step-by-Step (First Time)

### 1. Prerequisites
Ensure you have Node.js installed:
```bash
node --version  # Should show v18.0.0 or higher
```

If not installed, get it from: https://nodejs.org/

### 2. Install Dependencies
```bash
npm install
```

This installs:
- `docx` v8.5.0 - Word document creation library
- `adm-zip` v0.5.10 - ZIP manipulation library

### 3. Build the Document
```bash
node build_final_local.js
```

**Expected output:**
```
Done - file written to: ./outputs/Saliev_Thermoelasticity_IMRAD.docx
```

### 4. Open the Document
- **Windows**: Double-click the .docx file or open in Microsoft Word
- **Mac**: Open with Pages or Microsoft Word
- **Linux**: Use LibreOffice Writer (7.0+)

## For Windows Users

### Using Command Prompt
```cmd
npm install
node build_final_local.js
```

### Using PowerShell
```powershell
npm install
node build_final_local.js
```

## For Linux/Mac Users

### Using the Automated Script
```bash
chmod +x setup_and_build.sh
./setup_and_build.sh
```

### Manual Commands
```bash
npm install
node build_final_local.js
```

## Document Structure

```
Title: НЕСТАЦИОНАРНАЯ ЗАДАЧА СВЯЗАННОЙ ТЕРМОУПРУГОСТИ...
Author: Салиев А.А.

📄 Sections:
├── Аннотация (Abstract)
├── 1. ВВЕДЕНИЕ (Introduction)
├── 2. ИСХОДНЫЕ СООТНОШЕНИЯ (Materials & Methods)
│   ├── 2.1. Обобщённый обратный закон Гука
│   ├── 2.2. Динамические уравнения движения
│   └── 2.3. Гиперболическое уравнение теплопереноса
├── 3. РЕЗУЛЬТАТЫ (Results)
│   ├── 3.1. Термически скорректированный тензор ⭐ Новизна 1
│   ├── 3.2. Уравнение баланса энергии ⭐ Новизна 2
│   ├── 3.3. Обобщение Бельтрами–Митчелла ⭐ Новизна 3
│   ├── 3.4. Запись через инженерные константы
│   └── 3.5. Граничные и начальные условия
├── 4. ОБСУЖДЕНИЕ (Discussion)
├── 5. ЗАКЛЮЧЕНИЕ (Conclusion)
└── СПИСОК ЛИТЕРАТУРЫ (References - 8 items)
```

## Equations Included

1. **ε^i_j** - Generalized Hooke's law with thermal expansion
2. **I₁** - First invariant of stress tensor
3. **∇²σ^i_j** - Dynamic Beltrami-Michell in curvilinear coordinates
4. **c₀Ṫ** - Hyperbolic heat conduction with relaxation time τ_T
5. **σ*^i_j** - Thermally corrected stress tensor (Novel)
6. **dU/dt** - Energy balance equation (Novel)
7. **∇²σ*^i_j** - Main result with finite heat speed (Novel)
8. **ε^i_j = (1+ν)/E σ^i_j** - Engineering constants form

All equations use subscripts, superscripts, fractions, Greek letters, and mathematical operators properly formatted in Word's equation editor format (OMML).

## Troubleshooting

### Problem: `npm: command not found`
**Solution:** Install Node.js from https://nodejs.org/ (npm comes with it)

### Problem: `Cannot find module 'docx'`
**Solution:** Run `npm install` first

### Problem: Permission denied
**Solution:** 
- Linux/Mac: `sudo npm install` or fix npm permissions
- Windows: Run Command Prompt as Administrator

### Problem: Output directory doesn't exist
**Solution:** The script automatically creates `./outputs/` directory

### Problem: Math equations look wrong
**Solution:** 
- Open in Microsoft Word for best results
- LibreOffice 7.0+ also supports OMML
- Google Docs does NOT fully support OMML

## Performance

- Installation time: ~30-60 seconds
- Build time: ~2-3 seconds
- Output file size: ~35-40 KB

## What's IMRAD?

**IMRAD** is a standard format for scientific papers:
- **I**ntroduction - Background, gap, objectives
- **M**aterials/Methods - How the research was done
- **R**esults - What was found (3 novel contributions)
- **A**nd
- **D**iscussion - Interpretation, comparison, implications

## Next Steps

After generating the document:
1. ✅ Open in Microsoft Word
2. ✅ Review all sections
3. ✅ Verify equations render correctly
4. ✅ Check formatting (fonts, spacing, alignment)
5. ✅ Save/export to PDF if needed

## Need Help?

1. Check `USAGE_INSTRUCTIONS.md` for detailed information
2. Check `README.md` for technical details
3. Ensure Node.js v18+ is installed
4. Ensure npm has internet access for package installation

## License & Attribution

The content represents research by Салиев А.А. (Saliev A.A.) from Tashkent State Economic University on coupled thermoelasticity with finite heat wave speed.
