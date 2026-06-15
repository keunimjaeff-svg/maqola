# Word Document Builder for Scientific Article

This package creates a Word document following IMRAD format for the scientific article on thermoelasticity.

## Requirements

- Node.js (v18 or higher)
- npm

## Installation

```bash
npm install
```

This will install:
- `docx` - For creating Word documents
- `adm-zip` - For post-processing to inject OMML math equations

## Usage

```bash
npm run build
```

Or directly:

```bash
node build_final.js
```

## Output

The script will generate: `Saliev_Thermoelasticity_IMRAD.docx`

This document contains:
- Title and author information
- Abstract with keywords
- Introduction section with literature review
- Materials and Methods section with mathematical formulations
- Results section with 3 novel contributions
- Discussion section comparing with existing work
- Conclusion section with numbered achievements
- References section

All mathematical equations are rendered as proper OMML (Office Math Markup Language) for native Word equation support.

## Document Structure (IMRAD)

- **I**ntroduction (Введение)
- **M**aterials and Methods (Исходные соотношения и методология)
- **R**esults (Результаты)
- **A**nd
- **D**iscussion (Обсуждение + Заключение)

## Mathematical Equations

The document includes 8 key equations:
1. Generalized Hooke's law (inverse form)
2. First invariant of stress tensor
3. Dynamic Beltrami-Michell equations
4. Hyperbolic heat conduction equation
5. Thermally corrected stress tensor
6. Energy balance equation
7. Main result: Generalized Beltrami-Michell with finite heat speed
8. Engineering constants form

All equations use proper mathematical notation with subscripts, superscripts, fractions, and special symbols.
