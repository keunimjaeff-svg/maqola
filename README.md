# Word Document Builder for Scientific Articles

This repository contains generators for two scientific articles on thermoelasticity by Saliev A.A. (Tashkent State Economic University).

## 📚 Available Articles

### Article 1: Theoretical Formulation (Теоретическая постановка)
**Script:** `build_final_local.js`  
**Output:** `outputs/Saliev_Thermoelasticity_IMRAD.docx`  
**Topic:** Нестационарная задача связанной термоупругости в напряжениях с учётом конечной скорости тепловых потоков в произвольной криволинейной системе координат

**Key Contributions:**
- Thermally corrected stress tensor
- Energy balance equation in stresses
- Generalized Beltrami-Michell equations with hyperbolic heat transfer

### Article 2: Numerical Methods (Численные методы)
**Script:** `build_article2.js`  
**Output:** `outputs/Saliev_Numerical_FDM_IMRAD.docx`  
**Topic:** Численное решение нестационарной задачи связанной термоупругости методом конечных разностей в цилиндрических координатах

**Key Contributions:**
- Finite difference scheme for cylindrical coordinates
- CFL stability condition for 3D thermoelasticity
- Coupled iterative algorithm with error estimation

## 🚀 Quick Start

```bash
# Install dependencies
npm install

# Generate Article 1
node build_final_local.js

# Generate Article 2
node build_article2.js

# Generate both articles
node build_final_local.js && node build_article2.js
```

## 📂 Output

All generated documents are saved in the `outputs/` directory:
- `Saliev_Thermoelasticity_IMRAD.docx` - Article 1
- `Saliev_Numerical_FDM_IMRAD.docx` - Article 2

## 📋 Requirements

- **Node.js** v18 or higher
- **npm** (comes with Node.js)
- **Microsoft Word** for best equation rendering (or LibreOffice 7.0+)

## 📦 Dependencies

```json
{
  "docx": "^8.5.0",     // Word document generation
  "adm-zip": "^0.5.10"  // ZIP manipulation for OMML injection
}
```

## 📖 Documentation

- **QUICKSTART.md** - Quick start guide in Russian
- **USAGE_INSTRUCTIONS.md** - Detailed usage instructions
- **README_ARTICLE2.md** - Detailed info about Article 2

## 🔬 Document Structure (IMRAD Format)

Both articles follow the IMRAD structure:

1. **Introduction (Введение)** - Background, literature review, objectives
2. **Materials and Methods (Методология)** - Mathematical formulations
3. **Results (Результаты)** - Main findings (3 novel contributions each)
4. **Discussion (Обсуждение)** - Interpretation, comparison, applications
5. **Conclusion (Заключение)** - Summary of achievements

## 📐 Mathematical Content

### Article 1: 8 Equations
1. Generalized Hooke's law (inverse form)
2. First invariant of stress tensor
3. Dynamic Beltrami-Michell equations
4. Hyperbolic heat conduction equation
5. Thermally corrected stress tensor
6. Energy balance equation
7. Main result with finite heat speed
8. Engineering constants form

### Article 2: 8 Equations
1. Heat equation in cylindrical coordinates
2. Stress equilibrium (radial direction)
3. Explicit finite difference scheme
4. CFL stability condition
5. Discretized stress-strain relation
6. Strain from displacement
7. Coupled iteration scheme
8. L² error estimate

All equations use OMML (Office Math Markup Language) for native Word rendering.

## 🎯 How It Works

1. **Stage 1:** Generate .docx with placeholders using `docx` library
2. **Stage 2:** Create OMML (Office Math ML) XML for each equation
3. **Stage 3:** Unzip .docx file (it's a ZIP archive)
4. **Stage 4:** Replace placeholders with real OMML math in `word/document.xml`
5. **Stage 5:** Rezip and save final document

## 📝 Page Layout

- **Paper:** A4 (210 × 297 mm)
- **Font:** Times New Roman, 12pt
- **Line spacing:** 1.5 (240 twips)
- **Alignment:** Justified
- **Margins:** Standard academic format

## 🔗 Article Relationship

| Aspect | Article 1 | Article 2 |
|--------|-----------|-----------|
| **Focus** | Theoretical formulation | Numerical solution |
| **Approach** | Analytical | Computational |
| **Coordinates** | Arbitrary curvilinear | Cylindrical (r,φ,z) |
| **Method** | Beltrami-Michell equations | Finite difference method |
| **Result** | New system of equations | Algorithm & verification |

## 🛠️ Troubleshooting

### Error: `Cannot find module 'docx'`
**Solution:** Run `npm install` first

### Math equations not rendering
**Solution:** Open in Microsoft Word 2010+ (not Google Docs)

### Russian characters incorrect
**Solution:** Ensure UTF-8 encoding

## 📄 License

Academic articles by Saliev A.A., Tashkent State Economic University.

## 👤 Author

**Салиев А.А. (Saliev A.A.)**  
Ташкентский государственный экономический университет  
Кафедра высшей и прикладной математики  
г. Ташкент, Узбекистан
