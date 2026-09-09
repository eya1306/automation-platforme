"""
c_parser_core.py
================
Analyseur syntaxique de code C embarqué.
Extrait les variables, types, conditions et définitions
d'une fonction donnée dans des fichiers .c / .h
"""

import re
import os
from dataclasses import dataclass, field
from typing import List, Optional, Tuple


# ---------------------------------------------------------------------------
# Structures de données
# ---------------------------------------------------------------------------

@dataclass
class Variable:
    name: str
    type_: str
    initial_value: str = ""
    scope: str = "local"         # "local" | "param" | "global"
    line_number: int = 0
    is_pointer: bool = False
    is_array: bool = False
    array_size: str = ""
    is_const: bool = False
    is_volatile: bool = False
    is_static: bool = False
    comment: str = ""


@dataclass
class Condition:
    kind: str            # "if" | "else if" | "else" | "switch" | "case" | "for" | "while" | "do-while"
    expression: str
    line_number: int
    nesting_level: int = 0


@dataclass
class Definition:
    kind: str            # "define" | "typedef" | "enum" | "struct" | "union"
    name: str
    value: str
    line_number: int
    file_source: str = ""


@dataclass
class FunctionCall:
    callee: str
    arguments: str
    line_number: int


@dataclass
class FunctionInfo:
    name: str
    return_type: str = ""
    parameters: List[Variable] = field(default_factory=list)
    local_variables: List[Variable] = field(default_factory=list)
    global_variables: List[Variable] = field(default_factory=list)
    conditions: List[Condition] = field(default_factory=list)
    definitions: List[Definition] = field(default_factory=list)
    function_calls: List[FunctionCall] = field(default_factory=list)
    source_file: str = ""
    start_line: int = 0
    end_line: int = 0
    body: str = ""
    found: bool = False


# ---------------------------------------------------------------------------
# Utilitaires
# ---------------------------------------------------------------------------

# Qualificateurs C courants en embarqué
QUALIFIERS = r'(?:(?:const|volatile|static|extern|inline|register|unsigned|signed|'  \
             r'__volatile__|__const__|__attribute__\s*\(.*?\)|'                        \
             r'CONST|VOLATILE|STATIC)\s+)*'

# Types C de base + types embarqués
BASE_TYPES = (
    r'(?:unsigned\s+|signed\s+)?'
    r'(?:long\s+long|long\s+double|long|short|'
    r'int8_t|int16_t|int32_t|int64_t|'
    r'uint8_t|uint16_t|uint32_t|uint64_t|'
    r'float32_t|float64_t|'
    r'bool|_Bool|'
    r'int|char|float|double|void|'
    r'[A-Z_][A-Z0-9_]*_t|'          # typedefs style AUTOSAR/MISRA
    r'[A-Za-z_][A-Za-z0-9_]*)'      # tout autre typedef
)

VAR_DECL_RE = re.compile(
    rf'^[ \t]*{QUALIFIERS}({BASE_TYPES})'           # type
    r'(\s*\*+\s*|\s+)'                              # pointeur ou espace
    r'([A-Za-z_][A-Za-z0-9_]*)'                    # nom variable
    r'(\[([^\]]*)\])?'                              # tableau optionnel
    r'\s*(?:=\s*([^;,\n]+?))?'                     # valeur initiale optionnelle
    r'\s*;',
    re.MULTILINE
)

PARAM_RE = re.compile(
    rf'({QUALIFIERS}{BASE_TYPES})'                  # type (avec qualificateurs)
    r'(\s*\*+\s*|\s+)'                              # pointeur ou espace
    r'([A-Za-z_][A-Za-z0-9_]*)'                    # nom param
    r'(\[([^\]]*)\])?'                              # tableau optionnel
)

CONDITION_RE = re.compile(
    r'\b(if|else\s+if|else|switch|case|for|while|do)\b'
    r'(?:\s*\(([^)]*(?:\([^)]*\)[^)]*)*)\))?',
    re.MULTILINE
)

DEFINE_RE = re.compile(r'^\s*#\s*define\s+([A-Za-z_][A-Za-z0-9_]*)\s+(.*?)(?://.*)?$', re.MULTILINE)
TYPEDEF_RE = re.compile(r'\btypedef\b(.+?)\b([A-Za-z_][A-Za-z0-9_]*)\s*;', re.DOTALL)
ENUM_RE    = re.compile(r'\benum\s+([A-Za-z_][A-Za-z0-9_]*)\s*\{([^}]*)\}', re.DOTALL)
STRUCT_RE  = re.compile(r'\b(struct|union)\s+([A-Za-z_][A-Za-z0-9_]*)\s*\{([^}]*)\}', re.DOTALL)

FUNC_CALL_RE = re.compile(
    r'\b([A-Za-z_][A-Za-z0-9_]*)\s*\(([^;{]*?)\)',
    re.MULTILINE
)

COMMENT_LINE_RE  = re.compile(r'//.*?$', re.MULTILINE)
COMMENT_BLOCK_RE = re.compile(r'/\*.*?\*/', re.DOTALL)

C_KEYWORDS = {
    'if','else','switch','case','for','while','do','break','continue',
    'return','goto','sizeof','typedef','struct','union','enum',
    'const','volatile','static','extern','inline','register',
    'void','int','char','float','double','short','long','unsigned','signed',
    'default','NULL','true','false'
}


def strip_comments(source: str) -> Tuple[str, str]:
    """Retourne (source_sans_comments, version_avec_numéros_de_ligne_préservés)."""
    # Version "propre" pour la regex (les commentaires sont remplacés par des espaces
    # pour conserver la numérotation de lignes)
    clean = COMMENT_BLOCK_RE.sub(lambda m: '\n' * m.group(0).count('\n'), source)
    clean = COMMENT_LINE_RE.sub('', clean)
    return clean, source


def extract_inline_comment(line: str) -> str:
    """Extrait un commentaire inline `// ...` s'il existe."""
    m = re.search(r'//\s*(.*)', line)
    return m.group(1).strip() if m else ""


# ---------------------------------------------------------------------------
# Extraction du corps d'une fonction
# ---------------------------------------------------------------------------

def find_function_body(source_clean: str, func_name: str) -> Optional[Tuple[str, int, int, str]]:
    """
    Localise le corps d'une fonction dans le source nettoyé.
    Retourne (corps_brut, ligne_debut, ligne_fin, signature_retour_type)
    ou None si non trouvée.
    """
    # Regex pour trouver la signature de la fonction
    # Supporte les fonctions multilignes et les attributs embarqués
    pattern = re.compile(
        rf'((?:(?:static|inline|extern|const|volatile|'
        rf'__attribute__\s*\([^)]*\)\s*|STATIC_FUNC\s*\([^)]*\)\s*|'
        rf'[A-Za-z_][A-Za-z0-9_]*\s+)*)'
        rf'(?:[A-Za-z_][A-Za-z0-9_\s\*]*\s+)?\*?\s*)'  # return type
        rf'\b{re.escape(func_name)}\s*\('                # function name
        rf'([^)]*(?:\([^)]*\)[^)]*)*)\)',               # parameters
        re.MULTILINE
    )

    match = pattern.search(source_clean)
    if not match:
        return None

    # Chercher l'accolade ouvrante
    brace_start = source_clean.find('{', match.end())
    if brace_start == -1:
        # Peut-être une déclaration sans corps (prototype)
        return None

    # Vérifier qu'il n'y a que des espaces/commentaires entre ) et {
    between = source_clean[match.end():brace_start].strip()
    if between and not re.match(r'^[\s]*$', between):
        # Chercher le prochain { après une éventuelle liste d'initialiseurs
        pass

    # Compter les accolades pour trouver la fin
    depth = 0
    brace_end = brace_start
    for i, ch in enumerate(source_clean[brace_start:], start=brace_start):
        if ch == '{':
            depth += 1
        elif ch == '}':
            depth -= 1
            if depth == 0:
                brace_end = i
                break

    body = source_clean[brace_start + 1: brace_end]

    # Calculer les numéros de ligne
    start_line = source_clean[:match.start()].count('\n') + 1
    end_line   = source_clean[:brace_end].count('\n') + 1

    # Extraire le type de retour (approximation)
    sig = match.group(0)
    ret_type_match = re.match(
        rf'(.+?)\s*\*?\s*\b{re.escape(func_name)}\s*\(', sig, re.DOTALL
    )
    ret_type = ret_type_match.group(1).strip() if ret_type_match else "unknown"
    ret_type = re.sub(r'\s+', ' ', ret_type)

    return body, start_line, end_line, ret_type, match.group(2)  # body, l1, l2, ret, params_str


# ---------------------------------------------------------------------------
# Extraction des paramètres
# ---------------------------------------------------------------------------

def parse_parameters(params_str: str, start_line: int) -> List[Variable]:
    params = []
    if not params_str or params_str.strip() in ('', 'void'):
        return params

    # Séparer les paramètres (attention aux virgules dans les tableaux)
    parts = re.split(r',(?![^\[]*\])', params_str)
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m = PARAM_RE.match(part)
        if m:
            qualifiers_and_type = (m.group(1) or "").strip()
            pointer_or_space    = (m.group(2) or "").strip()
            name                = m.group(3).strip()
            array_bracket       = m.group(4) or ""
            array_size          = m.group(5) or ""

            is_ptr   = '*' in pointer_or_space or '*' in part
            is_array = bool(array_bracket)
            is_const = 'const' in qualifiers_and_type.lower()
            is_vol   = 'volatile' in qualifiers_and_type.lower()

            # Nettoyer le type
            clean_type = re.sub(r'\b(const|volatile|static|extern)\b', '', qualifiers_and_type).strip()
            clean_type = re.sub(r'\s+', ' ', clean_type).strip()
            if is_ptr:
                clean_type += '*'

            params.append(Variable(
                name=name,
                type_=clean_type,
                scope="param",
                line_number=start_line,
                is_pointer=is_ptr,
                is_array=is_array,
                array_size=array_size,
                is_const=is_const,
                is_volatile=is_vol,
            ))
    return params


# ---------------------------------------------------------------------------
# Extraction des variables locales
# ---------------------------------------------------------------------------

def parse_local_variables(body: str, func_start_line: int, original_source: str) -> List[Variable]:
    variables = []
    seen_names = set()

    lines = body.split('\n')
    body_start_line = func_start_line + original_source[:original_source.find(body)].count('\n') if body in original_source else func_start_line

    for rel_line, line in enumerate(lines, start=1):
        abs_line = func_start_line + rel_line
        # Ignorer les lignes de préprocesseur
        stripped = line.strip()
        if stripped.startswith('#'):
            continue

        m = VAR_DECL_RE.match(line)
        if not m:
            continue

        type_full  = m.group(1).strip()
        name       = m.group(3).strip()
        array_size = m.group(5) or ""
        init_val   = (m.group(6) or "").strip()

        if name in C_KEYWORDS or name in seen_names:
            continue
        seen_names.add(name)

        is_ptr     = '*' in (m.group(2) or "")
        is_const   = 'const' in line[:m.start(3)]
        is_vol     = 'volatile' in line[:m.start(3)]
        is_static  = 'static' in line[:m.start(3)]
        is_array   = bool(array_size)
        comment    = extract_inline_comment(line)

        variables.append(Variable(
            name=name,
            type_=type_full + ('*' if is_ptr else ''),
            initial_value=init_val,
            scope="local",
            line_number=abs_line,
            is_pointer=is_ptr,
            is_array=is_array,
            array_size=array_size,
            is_const=is_const,
            is_volatile=is_vol,
            is_static=is_static,
            comment=comment,
        ))

    return variables


# ---------------------------------------------------------------------------
# Extraction des variables globales
# ---------------------------------------------------------------------------

def parse_global_variables(source_clean: str) -> List[Variable]:
    """
    Extrait les variables déclarées au niveau global
    (en dehors de tout corps de fonction).
    """
    variables = []
    seen_names: set = set()

    # Reconstruire le texte hors des accolades de niveau >= 1
    # (= tout ce qui est à l'intérieur d'une fonction / struct / enum)
    outside: list = []
    depth = 0
    for ch in source_clean:
        if ch == '{':
            depth += 1
            outside.append('\n')       # conserver la numérotation de lignes
        elif ch == '}':
            depth -= 1
            outside.append('\n')
        elif depth == 0:
            outside.append(ch)
        else:
            outside.append('\n' if ch == '\n' else ' ')

    global_source = ''.join(outside)

    lines = global_source.split('\n')
    for line_no, line in enumerate(lines, start=1):
        stripped = line.strip()
        # Ignorer lignes vides, préprocesseur, commentaires
        if not stripped or stripped.startswith('#') or stripped.startswith('//'):
            continue

        m = VAR_DECL_RE.match(line)
        if not m:
            continue

        name = m.group(3).strip()
        if name in C_KEYWORDS or name in seen_names:
            continue
        seen_names.add(name)

        type_full  = m.group(1).strip()
        array_size = m.group(5) or ""
        init_val   = (m.group(6) or "").strip()
        is_ptr     = '*' in (m.group(2) or "")
        is_const   = 'const'    in line[:m.start(3)]
        is_vol     = 'volatile' in line[:m.start(3)]
        is_static  = 'static'   in line[:m.start(3)]
        is_array   = bool(array_size)
        comment    = extract_inline_comment(line)

        variables.append(Variable(
            name=name,
            type_=type_full + ('*' if is_ptr else ''),
            initial_value=init_val,
            scope="global",
            line_number=line_no,
            is_pointer=is_ptr,
            is_array=is_array,
            array_size=array_size,
            is_const=is_const,
            is_volatile=is_vol,
            is_static=is_static,
            comment=comment,
        ))

    return variables


# ---------------------------------------------------------------------------
# Extraction des conditions
# ---------------------------------------------------------------------------

def parse_conditions(body: str, func_start_line: int) -> List[Condition]:
    conditions = []
    lines = body.split('\n')
    nesting = 0

    for rel_line, line in enumerate(lines, start=1):
        abs_line = func_start_line + rel_line
        # Compter les accolades pour le niveau d'imbrication
        nesting += line.count('{') - line.count('}')
        nesting = max(0, nesting)

        for m in CONDITION_RE.finditer(line):
            kind = m.group(1).strip()
            # Normaliser "else if"
            kind = re.sub(r'\s+', ' ', kind)
            expr = (m.group(2) or "").strip()

            conditions.append(Condition(
                kind=kind,
                expression=expr,
                line_number=abs_line,
                nesting_level=nesting,
            ))

    return conditions


# ---------------------------------------------------------------------------
# Extraction des définitions globales (#define, typedef, enum, struct)
# ---------------------------------------------------------------------------

def parse_definitions(source: str, file_path: str) -> List[Definition]:
    defs = []

    # #define
    for m in DEFINE_RE.finditer(source):
        line_no = source[:m.start()].count('\n') + 1
        defs.append(Definition(
            kind="define",
            name=m.group(1),
            value=m.group(2).strip(),
            line_number=line_no,
            file_source=file_path,
        ))

    # typedef
    for m in TYPEDEF_RE.finditer(source):
        line_no = source[:m.start()].count('\n') + 1
        defs.append(Definition(
            kind="typedef",
            name=m.group(2).strip(),
            value=m.group(1).strip().replace('\n', ' '),
            line_number=line_no,
            file_source=file_path,
        ))

    # enum
    for m in ENUM_RE.finditer(source):
        line_no = source[:m.start()].count('\n') + 1
        values = ', '.join(v.strip() for v in m.group(2).split(',') if v.strip())
        defs.append(Definition(
            kind="enum",
            name=m.group(1),
            value=values,
            line_number=line_no,
            file_source=file_path,
        ))

    # struct / union
    for m in STRUCT_RE.finditer(source):
        line_no = source[:m.start()].count('\n') + 1
        defs.append(Definition(
            kind=m.group(1),
            name=m.group(2),
            value=m.group(3).replace('\n', ' ').strip(),
            line_number=line_no,
            file_source=file_path,
        ))

    return defs


# ---------------------------------------------------------------------------
# Extraction des appels de fonctions dans le corps
# ---------------------------------------------------------------------------

def parse_function_calls(body: str, func_start_line: int, own_name: str) -> List[FunctionCall]:
    calls = []
    seen = set()
    lines = body.split('\n')

    for rel_line, line in enumerate(lines, start=1):
        abs_line = func_start_line + rel_line
        for m in FUNC_CALL_RE.finditer(line):
            callee = m.group(1)
            if callee in C_KEYWORDS or callee == own_name:
                continue
            args = m.group(2).strip()
            key = (callee, abs_line)
            if key not in seen:
                seen.add(key)
                calls.append(FunctionCall(callee=callee, arguments=args, line_number=abs_line))

    return calls


# ---------------------------------------------------------------------------
# Fonction principale d'analyse
# ---------------------------------------------------------------------------

def analyze_function(func_name: str, file_paths: List[str]) -> FunctionInfo:
    """
    Analyse une fonction C dans une liste de fichiers.
    Retourne un objet FunctionInfo complet.
    """
    info = FunctionInfo(name=func_name)
    all_definitions: List[Definition] = []
    all_globals: List[Variable] = []
    seen_global_names: set = set()

    for fpath in file_paths:
        if not os.path.isfile(fpath):
            continue
        with open(fpath, 'r', encoding='utf-8', errors='replace') as f:
            raw = f.read()

        clean_source, _ = strip_comments(raw)

        # Variables globales (tous les fichiers)
        for gvar in parse_global_variables(clean_source):
            if gvar.name not in seen_global_names:
                seen_global_names.add(gvar.name)
                all_globals.append(gvar)

        # Définitions globales (tous les fichiers)
        all_definitions.extend(parse_definitions(raw, fpath))

        if info.found:
            continue  # Déjà trouvée, on continue pour les globals/définitions

        result = find_function_body(clean_source, func_name)
        if result is None:
            continue

        body, start_line, end_line, ret_type, params_str = result
        info.found       = True
        info.source_file = fpath
        info.start_line  = start_line
        info.end_line    = end_line
        info.return_type = ret_type
        info.body        = body

        # Paramètres
        info.parameters = parse_parameters(params_str, start_line)

        # Variables locales
        info.local_variables = parse_local_variables(body, start_line, clean_source)

        # Conditions
        info.conditions = parse_conditions(body, start_line)

        # Appels de fonctions
        info.function_calls = parse_function_calls(body, start_line, func_name)

    info.definitions     = all_definitions
    info.global_variables = all_globals
    return info
