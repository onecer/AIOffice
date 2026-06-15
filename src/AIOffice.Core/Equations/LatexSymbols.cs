namespace AIOffice.Core.Equations;

/// <summary>
/// The literal-symbol vocabulary of the supported LaTeX subset: Greek letters,
/// binary/relational operators, arrows and named constants, each mapped to the
/// Unicode character OMML stores. These are commands that emit a single math text
/// run (no arguments); structural commands (\frac, \sqrt, …) are handled by the
/// parser, not this table.
/// </summary>
internal static class LatexSymbols
{
    /// <summary>command name (without backslash) → Unicode replacement.</summary>
    public static readonly IReadOnlyDictionary<string, string> Map = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        // Lowercase Greek
        ["alpha"] = "α",
        ["beta"] = "β",
        ["gamma"] = "γ",
        ["delta"] = "δ",
        ["epsilon"] = "ϵ",
        ["varepsilon"] = "ε",
        ["zeta"] = "ζ",
        ["eta"] = "η",
        ["theta"] = "θ",
        ["vartheta"] = "ϑ",
        ["iota"] = "ι",
        ["kappa"] = "κ",
        ["lambda"] = "λ",
        ["mu"] = "μ",
        ["nu"] = "ν",
        ["xi"] = "ξ",
        ["omicron"] = "ο",
        ["pi"] = "π",
        ["varpi"] = "ϖ",
        ["rho"] = "ρ",
        ["varrho"] = "ϱ",
        ["sigma"] = "σ",
        ["varsigma"] = "ς",
        ["tau"] = "τ",
        ["upsilon"] = "υ",
        ["phi"] = "ϕ",
        ["varphi"] = "φ",
        ["chi"] = "χ",
        ["psi"] = "ψ",
        ["omega"] = "ω",

        // Uppercase Greek
        ["Gamma"] = "Γ",
        ["Delta"] = "Δ",
        ["Theta"] = "Θ",
        ["Lambda"] = "Λ",
        ["Xi"] = "Ξ",
        ["Pi"] = "Π",
        ["Sigma"] = "Σ",
        ["Upsilon"] = "Υ",
        ["Phi"] = "Φ",
        ["Psi"] = "Ψ",
        ["Omega"] = "Ω",

        // Binary operators and relations
        ["times"] = "×",
        ["div"] = "÷",
        ["cdot"] = "⋅",
        ["pm"] = "±",
        ["mp"] = "∓",
        ["ast"] = "∗",
        ["star"] = "⋆",
        ["circ"] = "∘",
        ["bullet"] = "∙",
        ["leq"] = "≤",
        ["le"] = "≤",
        ["geq"] = "≥",
        ["ge"] = "≥",
        ["neq"] = "≠",
        ["ne"] = "≠",
        ["equiv"] = "≡",
        ["approx"] = "≈",
        ["cong"] = "≅",
        ["sim"] = "∼",
        ["simeq"] = "≃",
        ["propto"] = "∝",
        ["ll"] = "≪",
        ["gg"] = "≫",
        ["doteq"] = "≐",
        ["asymp"] = "≍",
        ["bowtie"] = "⋈",
        ["prec"] = "≺",
        ["succ"] = "≻",
        ["preceq"] = "⪯",
        ["succeq"] = "⪰",
        ["subsetneq"] = "⊊",
        ["supsetneq"] = "⊋",
        ["sqsubset"] = "⊏",
        ["sqsupset"] = "⊐",
        ["lesssim"] = "≲",
        ["gtrsim"] = "≳",
        ["leqslant"] = "⩽",
        ["geqslant"] = "⩾",
        ["triangleq"] = "≜",
        ["ncong"] = "≇",
        ["nsim"] = "≁",

        // Arrows
        ["to"] = "→",
        ["rightarrow"] = "→",
        ["leftarrow"] = "←",
        ["leftrightarrow"] = "↔",
        ["Rightarrow"] = "⇒",
        ["Leftarrow"] = "⇐",
        ["Leftrightarrow"] = "⇔",
        ["mapsto"] = "↦",
        ["longmapsto"] = "⟼",
        ["uparrow"] = "↑",
        ["downarrow"] = "↓",
        ["updownarrow"] = "↕",
        ["Uparrow"] = "⇑",
        ["Downarrow"] = "⇓",
        ["Updownarrow"] = "⇕",
        ["longrightarrow"] = "⟶",
        ["longleftarrow"] = "⟵",
        ["longleftrightarrow"] = "⟷",
        ["Longrightarrow"] = "⟹",
        ["Longleftarrow"] = "⟸",
        ["Longleftrightarrow"] = "⟺",
        ["hookrightarrow"] = "↪",
        ["hookleftarrow"] = "↩",
        ["rightharpoonup"] = "⇀",
        ["rightharpoondown"] = "⇁",
        ["leftharpoonup"] = "↼",
        ["leftharpoondown"] = "↽",
        ["rightleftharpoons"] = "⇌",
        ["nrightarrow"] = "↛",
        ["nleftarrow"] = "↚",
        ["nleftrightarrow"] = "↮",
        ["nRightarrow"] = "⇏",
        ["nLeftarrow"] = "⇍",
        ["nLeftrightarrow"] = "⇎",
        ["implies"] = "⇒",
        ["impliedby"] = "⇐",
        ["iff"] = "⟺",

        // Set theory and logic
        ["in"] = "∈",
        ["notin"] = "∉",
        ["ni"] = "∋",
        ["subset"] = "⊂",
        ["supset"] = "⊃",
        ["subseteq"] = "⊆",
        ["supseteq"] = "⊇",
        ["cup"] = "∪",
        ["cap"] = "∩",
        ["setminus"] = "∖",
        ["emptyset"] = "∅",
        ["forall"] = "∀",
        ["exists"] = "∃",
        ["nexists"] = "∄",
        ["neg"] = "¬",
        ["lnot"] = "¬",
        ["land"] = "∧",
        ["wedge"] = "∧",
        ["lor"] = "∨",
        ["vee"] = "∨",
        ["complement"] = "∁",
        ["mid"] = "∣",
        ["nmid"] = "∤",
        ["top"] = "⊤",
        ["bot"] = "⊥",
        ["therefore"] = "∴",
        ["because"] = "∵",
        ["oplus"] = "⊕",
        ["otimes"] = "⊗",
        ["perp"] = "⊥",
        ["parallel"] = "∥",
        ["nparallel"] = "∦",
        ["models"] = "⊨",
        ["vdash"] = "⊢",
        ["dashv"] = "⊣",
        ["sqsubseteq"] = "⊑",
        ["sqsupseteq"] = "⊒",
        ["bigcirc"] = "◯",
        ["odot"] = "⊙",
        ["ominus"] = "⊖",
        ["oslash"] = "⊘",
        ["uplus"] = "⊎",
        ["sqcup"] = "⊔",
        ["sqcap"] = "⊓",
        ["amalg"] = "⨿",
        ["dagger"] = "†",
        ["ddagger"] = "‡",

        // Big operators (used standalone, not as \sum_..^.. n-ary which the parser handles)
        ["partial"] = "∂",
        ["nabla"] = "∇",
        ["infty"] = "∞",
        ["aleph"] = "ℵ",
        ["hbar"] = "ℏ",
        ["ell"] = "ℓ",
        ["Re"] = "ℜ",
        ["Im"] = "ℑ",
        ["wp"] = "℘",
        ["angle"] = "∠",
        ["triangle"] = "△",
        ["square"] = "□",
        ["diamond"] = "⋄",

        // Dots
        ["ldots"] = "…",
        ["cdots"] = "⋯",
        ["vdots"] = "⋮",
        ["ddots"] = "⋱",
        ["dots"] = "…",

        // Delimiters as standalone symbols
        ["langle"] = "⟨",
        ["rangle"] = "⟩",
        ["lceil"] = "⌈",
        ["rceil"] = "⌉",
        ["lfloor"] = "⌊",
        ["rfloor"] = "⌋",
        ["|"] = "‖",
        ["backslash"] = "\\",

        // Misc
        ["prime"] = "′",
        ["degree"] = "°",
    };

    /// <summary>
    /// The "big operator" commands that become an n-ary OMML object whose
    /// glyph/grow behavior differs from a plain symbol. Each maps to the operator
    /// glyph Word uses inside <c>m:naryPr/m:chr</c>.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, string> NaryOperators = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["sum"] = "∑",
        ["prod"] = "∏",
        ["coprod"] = "∐",
        ["int"] = "∫",
        ["iint"] = "∬",
        ["iiint"] = "∭",
        ["oint"] = "∮",
        ["bigcup"] = "⋃",
        ["bigcap"] = "⋂",
        ["bigoplus"] = "⨁",
        ["bigotimes"] = "⨂",
        ["bigvee"] = "⋁",
        ["bigwedge"] = "⋀",
    };

    /// <summary>
    /// Accent commands (<c>\hat</c>, <c>\bar</c>, …) mapped to their combining /
    /// overscript glyph for <c>m:acc/m:accPr/m:chr</c>. <c>\bar</c> and
    /// <c>\overline</c> emit a bar object rather than an accent in Word, handled
    /// separately by the parser.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, string> Accents = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["hat"] = "̂",
        ["widehat"] = "̂",
        ["tilde"] = "̃",
        ["widetilde"] = "̃",
        ["vec"] = "⃗",
        ["dot"] = "̇",
        ["ddot"] = "̈",
        ["acute"] = "́",
        ["grave"] = "̀",
        ["check"] = "̌",
        ["breve"] = "̆",
    };

    /// <summary>The delimiter pairs <c>\left … \right</c> and bare brackets understand.</summary>
    public static readonly IReadOnlyDictionary<string, string> OpenDelimiters = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        ["("] = "(",
        ["["] = "[",
        ["\\{"] = "{",
        ["lbrace"] = "{",
        ["langle"] = "⟨",
        ["lceil"] = "⌈",
        ["lfloor"] = "⌊",
        ["lvert"] = "|",
        ["lVert"] = "‖",
        ["Vert"] = "‖",
        ["vert"] = "|",
        ["|"] = "|",
        ["."] = string.Empty, // \left. is an invisible delimiter
    };

    public static readonly IReadOnlyDictionary<string, string> CloseDelimiters = new Dictionary<string, string>(StringComparer.Ordinal)
    {
        [")"] = ")",
        ["]"] = "]",
        ["\\}"] = "}",
        ["rbrace"] = "}",
        ["rangle"] = "⟩",
        ["rceil"] = "⌉",
        ["rfloor"] = "⌋",
        ["rvert"] = "|",
        ["rVert"] = "‖",
        ["Vert"] = "‖",
        ["vert"] = "|",
        ["|"] = "|",
        ["."] = string.Empty,
    };

    /// <summary>Named functions that render as upright text operators (<c>\sin</c>, <c>\log</c>, <c>\lim</c>).</summary>
    public static readonly IReadOnlySet<string> Functions = new HashSet<string>(StringComparer.Ordinal)
    {
        "sin", "cos", "tan", "cot", "sec", "csc",
        "sinh", "cosh", "tanh", "coth",
        "arcsin", "arccos", "arctan",
        "exp", "log", "ln", "lg",
        "lim", "limsup", "liminf", "max", "min", "sup", "inf",
        "det", "dim", "ker", "deg", "gcd", "hom", "arg", "Pr",
        "mod", "bmod", "argmax", "argmin", "varlimsup", "varliminf", "injlim", "projlim",
    };

    /// <summary>
    /// The "limit-style" operators whose immediately following <c>_subscript</c>
    /// is rendered as a lower limit (under the operator), the TeX <c>\limits</c>
    /// convention for <c>\lim</c>, <c>\max</c>, <c>\sup</c>, …
    /// </summary>
    public static readonly IReadOnlySet<string> LowerLimitFunctions = new HashSet<string>(StringComparer.Ordinal)
    {
        "lim", "limsup", "liminf", "max", "min", "sup", "inf",
        "argmax", "argmin", "varlimsup", "varliminf", "injlim", "projlim",
    };
}
