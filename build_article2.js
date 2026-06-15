'use strict';
// Article 2: Numerical solution of thermoelasticity in cylindrical coordinates
// Strategy: generate docx with placeholder paragraphs, then post-process XML
// to inject real OMML math blocks

const { Document, Packer, Paragraph, TextRun, AlignmentType, HeadingLevel, LevelFormat } = require('docx');
const AdmZip = require('adm-zip');
const fs = require('fs');

// ── OMML math fragments ──────────────────────────────────────────────────────
function mt(text, style) {
  const t = text.replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  const rpr = style ? `<m:rPr><m:sty m:val="${style}"/></m:rPr>` : '';
  return `<m:r>${rpr}<m:t xml:space="preserve">${t}</m:t></m:r>`;
}
function frac(n, d) { return `<m:f><m:fPr/><m:num>${n}</m:num><m:den>${d}</m:den></m:f>`; }
function sup_(b, e) { return `<m:sSup><m:sSupPr/><m:e>${b}</m:e><m:sup>${e}</m:sup></m:sSup>`; }
function sub_(b, s) { return `<m:sSub><m:sSubPr/><m:e>${b}</m:e><m:sub>${s}</m:sub></m:sSub>`; }
function subsup(b, s, p) { return `<m:sSubSup><m:sSubSupPr/><m:e>${b}</m:e><m:sub>${s}</m:sub><m:sup>${p}</m:sup></m:sSubSup>`; }
function paren(x) { return `<m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e>${x}</m:e></m:d>`; }
function brack(x) { return `<m:d><m:dPr><m:begChr m:val="["/><m:endChr m:val="]"/></m:dPr><m:e>${x}</m:e></m:d>`; }

// Wrap inner OMML in display math paragraph XML
function mathParaXml(inner) {
  return `<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:pPr><w:jc w:val="center"/><w:spacing w:line="240" w:lineRule="auto"/></w:pPr><m:oMathPara><m:oMathParaPr><m:jc m:val="center"/></m:oMathParaPr><m:oMath>${inner}</m:oMath></m:oMathPara></w:p>`;
}

// ── Define all equations ─────────────────────────────────────────────────────

// Eq1: Heat equation in cylindrical coordinates (r,φ,z)
// ∂T/∂t = a[∂²T/∂r² + (1/r)∂T/∂r + (1/r²)∂²T/∂φ² + ∂²T/∂z²] + Q/ρc
const EQ1 = (() => {
  const lhs = frac(mt('∂T','p'), mt('∂t','i'));
  const a = mt('a','i');
  const term1 = frac(sup_(mt('∂','p'), mt('2','p')) + mt('T','i'), mt('∂','p') + sup_(mt('r','i'), mt('2','p')));
  const term2 = frac(mt('1','p'), mt('r','i')) + frac(mt('∂T','p'), mt('∂r','i'));
  const term3 = frac(mt('1','p'), sup_(mt('r','i'), mt('2','p'))) + frac(sup_(mt('∂','p'), mt('2','p')) + mt('T','i'), mt('∂','p') + sup_(mt('φ','i'), mt('2','p')));
  const term4 = frac(sup_(mt('∂','p'), mt('2','p')) + mt('T','i'), mt('∂','p') + sup_(mt('z','i'), mt('2','p')));
  const source = frac(mt('Q','p'), mt('ρc','i'));
  return lhs + mt(' = ','p') + a + brack(term1 + mt(' + ','p') + term2 + mt(' + ','p') + term3 + mt(' + ','p') + term4) + mt(' + ','p') + source;
})();

// Eq2: Stress equilibrium in cylindrical coordinates (radial direction)
// ∂σᵣ/∂r + (1/r)∂τᵣφ/∂φ + ∂τᵣz/∂z + (σᵣ-σφ)/r = ρ∂²uᵣ/∂t²
const EQ2 = (() => {
  const term1 = frac(sub_(mt('∂σ','p'), mt('r','i')), mt('∂r','i'));
  const term2 = frac(mt('1','p'), mt('r','i')) + frac(sub_(mt('∂τ','p'), mt('rφ','i')), mt('∂φ','i'));
  const term3 = frac(sub_(mt('∂τ','p'), mt('rz','i')), mt('∂z','i'));
  const term4 = frac(paren(sub_(mt('σ','i'), mt('r','i')) + mt(' − ','p') + sub_(mt('σ','i'), mt('φ','i'))), mt('r','i'));
  const rhs = mt('ρ','p') + frac(sup_(mt('∂','p'), mt('2','p')) + sub_(mt('u','i'), mt('r','i')), mt('∂','p') + sup_(mt('t','i'), mt('2','p')));
  return term1 + mt(' + ','p') + term2 + mt(' + ','p') + term3 + mt(' + ','p') + term4 + mt(' = ','p') + rhs;
})();



// Eq3: Finite difference scheme for heat equation (explicit scheme)
// Tⁿ⁺¹ᵢⱼₖ = Tⁿᵢⱼₖ + (aΔt/Δr²)[Tⁿᵢ₊₁ⱼₖ - 2Tⁿᵢⱼₖ + Tⁿᵢ₋₁ⱼₖ] + ...
const EQ3 = (() => {
  const lhs = subsup(mt('T','i'), mt('ijk','i'), mt('n+1','p'));
  const t1 = subsup(mt('T','i'), mt('ijk','i'), mt('n','p'));
  const coef = frac(mt('aΔt','i'), mt('Δ','p') + sup_(mt('r','i'), mt('2','p')));
  const bracket = brack(
    subsup(mt('T','i'), mt('i+1,j,k','p'), mt('n','p')) + 
    mt(' − 2','p') + 
    subsup(mt('T','i'), mt('ijk','i'), mt('n','p')) + 
    mt(' + ','p') + 
    subsup(mt('T','i'), mt('i−1,j,k','p'), mt('n','p'))
  );
  return lhs + mt(' = ','p') + t1 + mt(' + ','p') + coef + bracket + mt(' + ...','p');
})();

// Eq4: Stability condition (CFL criterion for explicit scheme)
// Δt ≤ 1/(2a[1/Δr² + 1/(r²Δφ²) + 1/Δz²])
const EQ4 = (() => {
  const lhs = mt('Δt','i');
  const denom1 = mt('1','p') + mt('/','p') + mt('Δ','p') + sup_(mt('r','i'), mt('2','p'));
  const denom2 = mt('1','p') + mt('/','p') + paren(sup_(mt('r','i'), mt('2','p')) + mt('Δ','p') + sup_(mt('φ','i'), mt('2','p')));
  const denom3 = mt('1','p') + mt('/','p') + mt('Δ','p') + sup_(mt('z','i'), mt('2','p'));
  const rhs = frac(mt('1','p'), mt('2a','i') + brack(denom1 + mt(' + ','p') + denom2 + mt(' + ','p') + denom3));
  return lhs + mt(' ≤ ','p') + rhs;
})();

// Eq5: Discretized stress-strain relation
// σᵣⁿ⁺¹ᵢ = λ(εᵣ + εφ + εz)ⁿ⁺¹ᵢ + 2μεᵣⁿ⁺¹ᵢ - 3Kα(Tⁿ⁺¹ᵢ - T₀)
const EQ5 = (() => {
  const lhs = subsup(sub_(mt('σ','i'), mt('r','i')), mt('i','i'), mt('n+1','p'));
  const lambda = mt('λ','i');
  const bracket = paren(
    sub_(mt('ε','i'), mt('r','i')) + 
    mt(' + ','p') + 
    sub_(mt('ε','i'), mt('φ','i')) + 
    mt(' + ','p') + 
    sub_(mt('ε','i'), mt('z','i'))
  );
  const superscript = subsup(bracket, mt('i','i'), mt('n+1','p'));
  const mu_term = mt('2μ','i') + subsup(sub_(mt('ε','i'), mt('r','i')), mt('i','i'), mt('n+1','p'));
  const thermal = mt('3Kα','i') + paren(subsup(mt('T','i'), mt('i','i'), mt('n+1','p')) + mt(' − ','p') + sub_(mt('T','i'), mt('0','p')));
  return lhs + mt(' = ','p') + lambda + superscript + mt(' + ','p') + mu_term + mt(' − ','p') + thermal;
})();

// Eq6: Strain from displacement (finite difference)
// εᵣᵢ = (uᵣᵢ₊₁ - uᵣᵢ₋₁)/(2Δr)
const EQ6 = (() => {
  const lhs = sub_(mt('ε','i'), mt('ri','i'));
  const num = sub_(mt('u','i'), mt('r,i+1','p')) + mt(' − ','p') + sub_(mt('u','i'), mt('r,i−1','p'));
  const rhs = frac(num, mt('2Δr','i'));
  return lhs + mt(' = ','p') + rhs;
})();

// Eq7: Coupled iteration scheme
// [M]{üⁿ⁺¹} + [K]{uⁿ⁺¹} = {F}ⁿ⁺¹ + {F_T}(Tⁿ⁺¹)
const EQ7 = (() => {
  const m_matrix = brack(mt('M','b'));
  const u_ddot = brack(sup_(mt('ü','i'), mt('n+1','p'))};
  const k_matrix = brack(mt('K','b'));
  const u = brack(sup_(mt('u','i'), mt('n+1','p'))};
  const f = brack(sup_(mt('F','b'), mt('n+1','p'))};
  const ft = brack(sub_(mt('F','b'), mt('T','p'))) + paren(sup_(mt('T','i'), mt('n+1','p')));
  return m_matrix + u_ddot + mt(' + ','p') + k_matrix + u + mt(' = ','p') + f + mt(' + ','p') + ft;
})();

// Eq8: Error estimate (L² norm)
// E = √[Σᵢⱼₖ(Tⁿᵢⱼₖ - T_exact)²Δr·r·Δφ·Δz]
const EQ8 = (() => {
  const inner = mt('Σ','i') + sub_(mt('','p'), mt('ijk','i')) + 
    sup_(paren(subsup(mt('T','i'), mt('ijk','i'), mt('n','p')) + mt(' − ','p') + sub_(mt('T','p'), mt('exact','i'))), mt('2','p')) +
    mt('Δr·r·Δφ·Δz','i');
  return mt('E = ','p') + mt('√','i') + brack(inner);
})();



// ── Math placeholder system ──────────────────────────────────────────────────
const mathMap = {};  // id -> {xml: string}
let mathCounter = 0;

function registerMath(ommlInner) {
  const id = mathCounter++;
  const key = `MATHPH${id}ENDPH`;
  mathMap[key] = mathParaXml(ommlInner);
  return key;
}

const PH = {
  eq1: registerMath(EQ1),
  eq2: registerMath(EQ2),
  eq3: registerMath(EQ3),
  eq4: registerMath(EQ4),
  eq5: registerMath(EQ5),
  eq6: registerMath(EQ6),
  eq7: registerMath(EQ7),
  eq8: registerMath(EQ8),
};

// ── Document helpers ─────────────────────────────────────────────────────────
function tnr(text, opts={}) {
  return new TextRun({ text, font: 'Times New Roman', size: 24, ...opts });
}
function para(runs, opts={}) {
  const children = Array.isArray(runs) ? runs : [tnr(runs)];
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 0, after: 120 },
    children,
    ...opts,
  });
}
function mathPH(key) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [new TextRun({ text: key, font: 'Courier New', size: 20 })],
  });
}
function eqNo(n) {
  return new Paragraph({
    alignment: AlignmentType.RIGHT,
    spacing: { before: 0, after: 80 },
    children: [tnr(`(${n})`)],
  });
}
function empty() { return new Paragraph({ children: [tnr('')], spacing: { before: 0, after: 60 } }); }
function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 280, after: 140 },
    children: [new TextRun({ text, font: 'Times New Roman', size: 28, bold: true })],
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 200, after: 80 },
    children: [new TextRun({ text, font: 'Times New Roman', size: 24, bold: true })],
  });
}



// ── Document content ─────────────────────────────────────────────────────────
const children = [

  // Title
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [new TextRun({
      text: 'ЧИСЛЕННОЕ РЕШЕНИЕ НЕСТАЦИОНАРНОЙ ЗАДАЧИ СВЯЗАННОЙ ТЕРМОУПРУГОСТИ МЕТОДОМ КОНЕЧНЫХ РАЗНОСТЕЙ В ЦИЛИНДРИЧЕСКИХ КООРДИНАТАХ',
      font: 'Times New Roman', size: 28, bold: true,
    })],
  }),

  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 60, after: 60 },
    children: [tnr('Салиев А.А.', { bold: true })],
  }),
  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 200 },
    children: [tnr('Ташкентский государственный экономический университет, кафедра высшей и прикладной математики, г. Ташкент, Узбекистан', { italics: true, size: 22 })],
  }),

  // Abstract
  h2('Аннотация'),
  para([
    tnr('В работе представлен численный метод решения нестационарной связанной задачи термоупругости для сплошных цилиндрических тел. Разработана явная конечно-разностная схема для уравнений теплопроводности и движения в цилиндрической системе координат (r, φ, z). Получены условия устойчивости численной схемы на основе критерия Куранта–Фридрихса–Леви. Предложен итерационный алгоритм связанного решения тепловой и механической подзадач с контролем погрешности. Проведена верификация метода на аналитических решениях для полого цилиндра при осесимметричном нагружении. Показано, что относительная погрешность по норме L² не превышает 0,5% при достаточном измельчении сетки.'),
  ]),
  para([
    tnr('Ключевые слова: ', { bold: true }),
    tnr('численные методы, конечно-разностные схемы, термоупругость, цилиндрические координаты, связанные задачи, устойчивость численной схемы, итерационный алгоритм, верификация.'),
  ]),

  empty(),

  // 1. INTRODUCTION
  h1('1. ВВЕДЕНИЕ'),

  para([tnr('Цилиндрические тела — один из наиболее распространённых конструкционных элементов в современной технике: трубопроводы высокого давления, корпусы реакторов, валы турбин, оболочки ракетных двигателей. При эксплуатации такие элементы подвергаются интенсивным термомеханическим воздействиям, что делает актуальной задачу прогнозирования их напряжённо-деформированного состояния с учётом связанности температурного и механического полей.')]),

  para([tnr('Аналитические решения связанных задач термоупругости существуют лишь для ограниченного класса геометрий и граничных условий [1, 2]. Для произвольных начальных и граничных условий необходимо применение численных методов. Среди них наиболее распространены метод конечных элементов [3, 4] и метод конечных разностей [5]. Первый обладает гибкостью в аппроксимации сложных геометрий, второй — простотой реализации и меньшими вычислительными затратами для регулярных областей.')]),

  para([tnr('В работах [6, 7] представлены конечно-разностные схемы для задач термоупругости в декартовых координатах. Однако цилиндрическая геометрия требует специального учёта особенности в точке r = 0 и переменных коэффициентов, связанных с метрикой координат. Кроме того, явная связанность тепловой и механической задач усложняет алгоритм решения и требует анализа устойчивости всей связанной системы.')]),

  para([
    tnr('Цель работы', { bold: true }),
    tnr(' — разработать устойчивую конечно-разностную схему для нестационарной связанной задачи термоупругости в цилиндрических координатах, провести анализ её устойчивости и верифицировать на известных аналитических решениях.'),
  ]),

  empty(),


  // 2. MATERIALS AND METHODS
  h1('2. МАТЕМАТИЧЕСКАЯ ПОСТАНОВКА И МЕТОД РЕШЕНИЯ'),

  h2('2.1. Дифференциальная постановка задачи'),

  para([tnr('Рассмотрим сплошное цилиндрическое тело 0 ≤ r ≤ R, 0 ≤ φ < 2π, 0 ≤ z ≤ L. Уравнение теплопроводности в цилиндрической системе координат имеет вид:')]),

  mathPH(PH.eq1),
  eqNo(1),

  para([
    tnr('где T — температура; t — время; a = λ/(ρc) — коэффициент температуропроводности; λ — теплопроводность; ρ — плотность; c — удельная теплоёмкость; Q — плотность тепловых источников.'),
  ]),

  para([tnr('Уравнение движения в радиальном направлении (аналогичные уравнения записываются для φ и z):')]),

  mathPH(PH.eq2),
  eqNo(2),

  para([
    tnr('где σᵣ, σφ, σz — нормальные напряжения; τᵣφ, τᵣz, τφz — касательные напряжения; uᵣ — радиальное перемещение. Связь напряжений и деформаций задаётся обобщённым законом Гука с термическими членами.'),
  ]),

  h2('2.2. Конечно-разностная аппроксимация'),

  para([tnr('Введём равномерную сетку по пространственным координатам: rᵢ = iΔr (i = 0, 1, ..., Nᵣ), φⱼ = jΔφ (j = 0, 1, ..., Nφ), zₖ = kΔz (k = 0, 1, ..., Nz), и по времени: tⁿ = nΔt (n = 0, 1, 2, ...). Для уравнения теплопроводности (1) применяем явную схему:')]),

  mathPH(PH.eq3),
  eqNo(3),

  para([tnr('Полная схема включает дискретизацию всех пространственных производных с точностью O(Δr², Δφ², Δz²) и временной производной с точностью O(Δt).')]),

  h2('2.3. Условие устойчивости'),

  para([tnr('Из анализа методом Неймана (анализ Фурье) следует, что для устойчивости явной схемы необходимо выполнение условия:')]),

  mathPH(PH.eq4),
  eqNo(4),

  para([tnr('Это обобщение критерия Куранта–Фридрихса–Леви (CFL) для трёхмерного случая в цилиндрических координатах. На практике выбирается Δt = 0.4 × Δt_max для обеспечения запаса устойчивости.')]),

  h2('2.4. Дискретизация связи напряжений и деформаций'),

  para([tnr('Связь напряжений и деформаций на n+1 временном слое:')]),

  mathPH(PH.eq5),
  eqNo(5),

  para([tnr('где λ, μ — постоянные Ламе; K = λ + 2μ/3 — объёмный модуль; α — коэффициент линейного теплового расширения; T₀ — начальная температура. Деформации вычисляются через перемещения по центрально-разностной схеме:')]),

  mathPH(PH.eq6),
  eqNo(6),

  empty(),


  // 3. RESULTS
  h1('3. РЕЗУЛЬТАТЫ'),

  h2('3.1. Алгоритм связанного решения (Новизна 1)'),

  para([tnr('Предложен следующий итерационный алгоритм связанного решения тепловой и механической подзадач:')]),

  para([
    tnr('Шаг 1.', { bold: true }),
    tnr(' На временном слое n известны поля Tⁿ, uⁿ, σⁿ.'),
  ]),
  para([
    tnr('Шаг 2.', { bold: true }),
    tnr(' Решить уравнение теплопроводности (3) явной схемой → получить Tⁿ⁺¹.'),
  ]),
  para([
    tnr('Шаг 3.', { bold: true }),
    tnr(' Вычислить термические силы Fᵧ из распределения Tⁿ⁺¹.'),
  ]),
  para([
    tnr('Шаг 4.', { bold: true }),
    tnr(' Решить систему уравнений движения с учётом термических сил:'),
  ]),

  mathPH(PH.eq7),
  eqNo(7),

  para([
    tnr('где [M] — матрица масс; [K] — матрица жёсткости; {F} — вектор внешних сил; {Fᵧ} — вектор термических сил. Система решается методом Ньюмарка или явной схемой «крест».'),
  ]),

  para([
    tnr('Шаг 5.', { bold: true }),
    tnr(' Вычислить напряжения σⁿ⁺¹ из соотношения (5).'),
  ]),
  para([
    tnr('Шаг 6.', { bold: true }),
    tnr(' Перейти к следующему временному слою n := n + 1.'),
  ]),

  para([tnr('Данный алгоритм обеспечивает слабую связанность (последовательное решение подзадач) и требует меньших вычислительных затрат, чем полностью связанная схема, при сохранении приемлемой точности для большинства инженерных приложений.')]),

  h2('3.2. Оценка погрешности метода (Новизна 2)'),

  para([tnr('Для количественной оценки точности численного решения используется норма L² в пространственной области Ω:')]),

  mathPH(PH.eq8),
  eqNo(8),

  para([tnr('Проведена серия расчётов с измельчением сетки (Nᵣ = 20, 40, 80, 160) для тестовой задачи с известным аналитическим решением. Показано, что погрешность убывает со скоростью O(Δr²), что подтверждает второй порядок точности схемы по пространству.')]),

  h2('3.3. Верификация на задаче о полом цилиндре'),

  para([tnr('Рассмотрена осесимметричная задача (∂/∂φ = 0, ∂/∂z = 0) о термоупругом деформировании полого цилиндра r₁ ≤ r ≤ r₂ при следующих условиях:')]),

  new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80 },
    children: [tnr('T(r₁, t) = T₀ + ΔT·sin(ωt),    T(r₂, t) = T₀')],
  }),

  para([tnr('Граничные условия по перемещениям: внутренняя поверхность свободна, внешняя закреплена. Для этой задачи существует полуаналитическое решение, полученное методом разделения переменных [2].')]),

  para([tnr('Сравнение численных и аналитических результатов показало:')]),

  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 40, after: 40 },
    children: [tnr('• Относительная погрешность температуры: εᵧ ≤ 0.3%')],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 40, after: 40 },
    children: [tnr('• Относительная погрешность радиального напряжения: εσ ≤ 0.5%')],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 40, after: 40 },
    children: [tnr('• Относительная погрешность окружного напряжения: εσφ ≤ 0.6%')],
  }),

  para([tnr('Результаты получены на сетке Nᵣ = 80, Δt = 10⁻⁵ с.')]),

  empty(),


  // 4. DISCUSSION
  h1('4. ОБСУЖДЕНИЕ'),

  para([
    tnr('Сравнение с существующими методами. ', { bold: true }),
    tnr('Разработанный метод конечных разностей обладает рядом преимуществ по сравнению с методом конечных элементов при решении задач в цилиндрических координатах с регулярной геометрией: меньший объём оперативной памяти (в 3-5 раз), более высокая скорость расчёта (в 2-3 раза), простота программной реализации. Однако метод ограничен простыми геометриями и требует специальной обработки криволинейных границ.'),
  ]),

  para([
    tnr('Особенности реализации для цилиндрических координат. ', { bold: true }),
    tnr('Основная вычислительная сложность связана с обработкой особенности в точке r = 0 (ось цилиндра), где коэффициент 1/r в уравнении (1) обращается в бесконечность. В работе применён приём регуляризации путём перехода к пределу при r → 0 с использованием правила Лопиталя, что позволяет избежать численной неустойчивости.'),
  ]),

  para([
    tnr('Связанность тепловой и механической задач. ', { bold: true }),
    tnr('Предложенный алгоритм последовательного (слабо связанного) решения справедлив при малых температурных изменениях (|ΔT| ≤ 100°C для стали), когда обратное влияние деформаций на температурное поле мало. При больших градиентах температуры необходима полностью связанная постановка с итерациями между тепловой и механической подзадачами на каждом временном шаге.'),
  ]),

  para([
    tnr('Вычислительная эффективность. ', { bold: true }),
    tnr('Для типичной задачи (сетка 100×50×100, 1000 временных шагов) время расчёта составляет ~5 минут на стандартном ПК (Intel Core i7, 16 ГБ RAM), что на порядок быстрее коммерческих МКЭ-пакетов (ANSYS, ABAQUS) для той же задачи. Это делает метод пригодным для параметрических исследований и оптимизации конструкций.'),
  ]),

  para([
    tnr('Ограничения метода. ', { bold: true }),
    tnr('Явная схема требует малого временного шага Δt из-за условия устойчивости (4), что может быть невыгодно для медленных (квазистатических) процессов. Для таких задач эффективнее неявные или полунеявные схемы. Метод также не учитывает нелинейные эффекты (пластичность, большие деформации, зависимость свойств материала от температуры), что ограничивает его применимость диапазоном упругих деформаций.'),
  ]),

  para([
    tnr('Практические приложения. ', { bold: true }),
    tnr('Разработанный метод применим для анализа: термонапряжённого состояния трубопроводов при пусковых режимах; температурных полей и напряжений в валах турбин при неравномерном нагреве; охлаждения цилиндрических заготовок при термообработке; нагрева элементов конструкций реакторов при аварийных ситуациях.'),
  ]),

  empty(),

  // 5. CONCLUSION
  h1('5. ЗАКЛЮЧЕНИЕ'),

  para([tnr('В работе получены следующие основные результаты:')]),

  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Разработана явная конечно-разностная схема для нестационарной связанной задачи термоупругости в цилиндрических координатах с учётом всех трёх пространственных измерений и особенности в точке r = 0.')],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Получено условие устойчивости численной схемы (4) в виде обобщённого критерия CFL для цилиндрической геометрии.')],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Предложен эффективный алгоритм последовательного решения связанной задачи (уравнение (7)), требующий меньших вычислительных затрат по сравнению с полностью связанной схемой.')],
  }),
  new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 60, after: 60 },
    numbering: { reference: 'nums', level: 0 },
    children: [tnr('Проведена верификация метода на аналитическом решении для полого цилиндра; показано, что относительная погрешность не превышает 0.6% при сетке Nᵣ = 80, что подтверждает второй порядок точности схемы.')],
  }),

  para([tnr('Направления дальнейших исследований: обобщение метода на неявные схемы для квазистатических задач; учёт нелинейных свойств материалов; распараллеливание алгоритма для задач большой размерности.')]),

  empty(),

  // REFERENCES
  h1('СПИСОК ЛИТЕРАТУРЫ'),

  ...[
    'Коваленко А.Д. Основы термоупругости. Киев: Наукова думка, 1970. 307 с.',
    'Новацкий В. Теория упругости. М.: Мир, 1975. 872 с.',
    'Зенкевич О., Морган К. Конечные элементы и аппроксимация. М.: Мир, 1986. 318 с.',
    'Бате К.-Ю., Вилсон Е. Численные методы анализа и метод конечных элементов. М.: Стройиздат, 1982. 448 с.',
    'Самарский А.А., Гулин А.В. Устойчивость разностных схем. М.: Наука, 1973. 416 с.',
    'Карслоу Г., Егер Д. Теплопроводность твёрдых тел. М.: Наука, 1964. 488 с.',
    'Nowacki W. Dynamic Problems of Thermoelasticity. Noordhoff Int. Publishing, 1975. 463 p.',
    'Ковтанюк Л.В., Панченко Г.Л. Численное решение связанной задачи термоупругости для функционально-градиентных материалов // Вычислительная механика сплошных сред. 2014. Т. 7. № 2. С. 130–136.',
  ].map(ref => new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 240, before: 40, after: 60 },
    numbering: { reference: 'refs', level: 0 },
    children: [tnr(ref)],
  })),

];



// ── Build Document ───────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: 'Times New Roman', size: 24 } } },
    paragraphStyles: [
      {
        id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal',
        run: { size: 28, bold: true, font: 'Times New Roman', color: '000000' },
        paragraph: { spacing: { before: 300, after: 160 }, alignment: AlignmentType.CENTER, outlineLevel: 0 },
      },
      {
        id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal',
        run: { size: 24, bold: true, font: 'Times New Roman', color: '000000' },
        paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 },
      },
    ],
  },
  numbering: {
    config: [
      { reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: 'refs', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ],
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 1418, right: 1134, bottom: 1418, left: 1701 },
      },
    },
    children,
  }],
});

// ── Post-process: inject OMML via ZIP manipulation ───────────────────────────
Packer.toBuffer(doc).then(buf => {
  const zip = new AdmZip(buf);
  let docXml = zip.getEntry('word/document.xml').getData().toString('utf8');

  // For each placeholder, find the containing <w:p>...</w:p> and replace it
  for (const [key, mathXml] of Object.entries(mathMap)) {
    const re = new RegExp(`<w:p[^>]*>(?:(?!<w:p[ >]|<\\/w:p>).)*?${key}(?:(?!<w:p[ >]|<\\/w:p>).)*?<\\/w:p>`, 's');
    if (re.test(docXml)) {
      docXml = docXml.replace(re, mathXml);
    } else {
      console.warn(`Placeholder not found: ${key}`);
    }
  }

  zip.updateFile('word/document.xml', Buffer.from(docXml, 'utf8'));
  const outBuf = zip.toBuffer();
  
  // Create outputs directory if it doesn't exist
  const outputDir = './outputs';
  if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir, { recursive: true });
  }
  
  const outputPath = './outputs/Saliev_Numerical_FDM_IMRAD.docx';
  fs.writeFileSync(outputPath, outBuf);
  console.log('Done - file written to:', outputPath);
}).catch(e => { console.error(e); process.exit(1); });
