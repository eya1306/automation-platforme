"""
Extracteur de call tree depuis code C/ASM embarqué - v2
Mode 1: Analyse locale, pas d'authentification

Corrections apportees par rapport a la v1 :
  1. Scan de .c, .h ET .s/.asm (pas seulement .c)
  2. Extensions insensibles a la casse (.c/.C, .h/.H, .s/.S, .asm/.ASM)
  3. Detection de fonction independante du type de retour
     (plus de liste figee de types -> gere unsigned long, BOOL,
     typedefs maison, macros/attributs avant le type, etc.)
  4. Parametres avec parentheses imbriquees geres via comptage de
     profondeur (callback function pointers, macros dans les args...)
  5. Fonctions definies en assembleur (.global / .globl + label)
     enregistrees pour resoudre le CSC/CSU des appels vers elles
"""

import re
import os
from pathlib import Path
from typing import List, Dict, Set, Tuple, Optional


class CCodeExtractor:
    """Extrait call tree depuis fichiers C/H/ASM"""

    _IMPL_FOLDERS = {'impl', 'src', 'source', 'sources', 'code', 'codes'}

    _EXCLUDED_CALLS = {
        'if', 'while', 'for', 'switch', 'else', 'do', 'return',
        'sizeof', 'typeof', 'offsetof', 'alignof',
        'printf', 'fprintf', 'sprintf', 'snprintf', 'sscanf', 'vprintf',
        'malloc', 'calloc', 'realloc', 'free',
        'memcpy', 'memset', 'memmove', 'memcmp',
        'strlen', 'strcpy', 'strncpy', 'strcmp', 'strncmp',
        'strcat', 'strncat', 'strchr', 'strstr',
        'assert', 'abort', 'exit',
        'fopen', 'fclose', 'fread', 'fwrite', 'fgets', 'fputs',
    }

    # Mots-clefs qui ne sont jamais des noms de fonctions/macros de def
    _CONTROL_KEYWORDS = {
        'if', 'while', 'for', 'switch', 'return', 'sizeof', 'else', 'do',
        'typeof', 'offsetof', 'alignof', 'catch', 'try'
    }

    # Macros courantes dans l'embarqué (OSEK, AUTOSAR, RTOS, ISR) définissant des fonctions
    _MACRO_FUNC_WRAPPERS = {
        'TASK', 'ISR', 'ISR2', 'ALARM', 'INTERRUPT_HANDLER',
        'THREAD_ENTRY', 'COMPONENT_TASK', 'EX_INTERRUPT_HANDLER'
    }

    # Extensions reconnues (en minuscules) par categorie
    _C_LIKE_EXT = {'.c', '.h'}
    _ASM_EXT = {'.s', '.asm'}

    def __init__(self, source_dir: str):
        self.source_dir = Path(source_dir)
        # {func_name: {'calls': [(callee, condition), ...], 'file': Path}}
        self.functions = {}

    # ------------------------------------------------------------------
    # API publique
    # ------------------------------------------------------------------

    def extract(self) -> List[List]:
        """Extrait et retourne les lignes pour le format Call Tree attendu"""

        print(f"Analyse {self.source_dir}...")

        all_files = self._collect_files()

        if not all_files:
            print("  [!] Aucun fichier .c/.h/.s trouve")
            return [["CSC", "CSU", "Calling function",
                     "CSC", "CSU", "Called function", "Condition"]]

        print(f"  {len(all_files)} fichiers detectes")

        for f in all_files:
            if f.suffix.lower() in self._C_LIKE_EXT:
                self._parse_c_file(f)
            elif f.suffix.lower() in self._ASM_EXT:
                self._parse_asm_file(f)

        rows = self._to_excel_rows()
        print(f"  OK - {len(rows)} lignes generees")

        return rows

    def extract_from_text(self, content: str) -> List[List]:
        """Extrait call tree directement depuis un texte de code C"""
        self.functions = {}
        self.parse_text(content, "pasted_code.c")
        return self._to_excel_rows()

    # ------------------------------------------------------------------
    # Collecte de fichiers (insensible a la casse d'extension)
    # ------------------------------------------------------------------

    def _collect_files(self) -> List[Path]:
        wanted = self._C_LIKE_EXT | self._ASM_EXT
        found = []
        for p in self.source_dir.rglob("*"):
            if p.is_file() and p.suffix.lower() in wanted:
                found.append(p)
        return sorted(found)

    # ------------------------------------------------------------------
    # Parsing fichiers C / H
    # ------------------------------------------------------------------

    def _parse_c_file(self, file_path: Path):
        """Parse un fichier C/H pour extraire fonctions et appels"""
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            self.parse_text(content, file_path.name, file_path)
        except Exception as e:
            print(f"  [!] Erreur {file_path.name}: {e}")

    def parse_text(self, content: str,
                   filename: str = "pasted_code.c",
                   file_path: Path = None):
        """Parse le contenu texte d'un code C directement.

        Detecte les definitions de fonctions independamment de leur
        type de retour : on cherche un identifiant suivi d'une liste
        de parametres entre parentheses EQUILIBREES puis d'une '{',
        au niveau top-level du fichier (profondeur d'accolades = 0).
        """
        try:
            if file_path is None:
                file_path = self.source_dir / filename

            clean = self._clean_code(content)
            n = len(clean)
            i = 0
            depth = 0

            while i < n:
                ch = clean[i]

                if ch == '{':
                    depth += 1
                    i += 1
                    continue
                if ch == '}':
                    depth = max(0, depth - 1)
                    i += 1
                    continue

                # On ne cherche des DEFINITIONS de fonctions qu'au
                # niveau fichier (depth == 0). A l'interieur d'une
                # fonction deja trouvee, on saute directement son
                # corps via le comptage d'accolades ci-dessus.
                if depth == 0 and (ch.isalpha() or ch == '_'):
                    ident_m = re.match(r'[a-zA-Z_][a-zA-Z0-9_]*(?:::[a-zA-Z_][a-zA-Z0-9_]*)*', clean[i:])
                    if ident_m:
                        ident = ident_m.group(0)
                        after = i + len(ident)
                        # on saute les espaces entre l'identifiant et '('
                        j = after
                        while j < n and clean[j] in ' \t\r\n':
                            j += 1
                        if (ident not in self._CONTROL_KEYWORDS
                                and j < n and clean[j] == '('):
                            close_paren = self._find_matching_paren(clean, j)
                            if close_paren != -1:
                                # Sauter les blancs, attributs, specificateurs (__attribute__, const, INTERRUPT...) entre ')' et '{'
                                k = close_paren + 1
                                while k < n:
                                    if clean[k] in ' \t\r\n':
                                        k += 1
                                    elif clean[k] == '(':
                                        matching = self._find_matching_paren(clean, k)
                                        if matching != -1:
                                            k = matching + 1
                                        else:
                                            break
                                    elif clean[k] in '{;=':
                                        break
                                    else:
                                        sub_ident_m = re.match(r'([a-zA-Z_][a-zA-Z0-9_]*)\s*\(', clean[k:])
                                        if sub_ident_m and sub_ident_m.group(1) not in {'__attribute__', '__attribute', '__declspec', 'alignas', 'typeof'}:
                                            break
                                        tok_m = re.match(r'[a-zA-Z0-9_]+', clean[k:])
                                        if tok_m:
                                            k += len(tok_m.group(0))
                                        else:
                                            k += 1

                                if k < n and clean[k] == '{':
                                    # -> C'est une definition de fonction
                                    func_name = ident
                                    if ident in self._MACRO_FUNC_WRAPPERS:
                                        inner = clean[j + 1:close_paren].strip()
                                        inner_m = re.match(r'([a-zA-Z_][a-zA-Z0-9_]*)', inner)
                                        if inner_m and inner_m.group(1) not in {
                                            'void', 'int', 'char', 'float', 'double', 'unsigned', 'signed'
                                        }:
                                            func_name = inner_m.group(1)

                                    body_start = k + 1
                                    body_end = self._find_matching_brace(
                                        clean, k)
                                    body = clean[body_start:body_end]

                                    calls_with_conditions = (
                                        self._extract_calls_with_conditions(
                                            body))

                                    self.functions[func_name] = {
                                        'calls': calls_with_conditions,
                                        'file': file_path
                                    }

                                    # on continue juste apres le corps
                                    i = body_end + 1
                                    continue
                        # Pas une definition -> avancer d'un identifiant
                        i = after
                        continue

                i += 1

        except Exception as e:
            print(f"  [!] Erreur parsing texte: {e}")

    # ------------------------------------------------------------------
    # Parsing fichiers assembleur (.s / .asm)
    # ------------------------------------------------------------------

    def _parse_asm_file(self, file_path: Path):
        """Enregistre les symboles globaux definis dans un fichier asm.

        On ne tente pas d'analyser les branchements conditionnels en
        assembleur (trop specifique au jeu d'instructions). On se
        contente d'enregistrer les fonctions/symboles exportes pour
        que les appels C -> ASM soient resolus (CSC/CSU non-nuls).
        """
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()

            # .global nom / .globl nom
            globals_found = set(re.findall(
                r'^\s*\.(?:global|globl)\s+([a-zA-Z_][a-zA-Z0-9_]*)',
                content, flags=re.MULTILINE))

            # Labels "nom:" en debut de ligne (fallback si pas de .global)
            labels_found = set(re.findall(
                r'^\s*([a-zA-Z_][a-zA-Z0-9_]*)\s*:',
                content, flags=re.MULTILINE))

            symbols = globals_found if globals_found else labels_found

            for sym in symbols:
                if sym not in self.functions:
                    self.functions[sym] = {
                        'calls': [],
                        'file': file_path
                    }
        except Exception as e:
            print(f"  [!] Erreur {file_path.name}: {e}")

    # ------------------------------------------------------------------
    # Utilitaires de nettoyage / parenthesage
    # ------------------------------------------------------------------

    @staticmethod
    def _clean_code(text: str) -> str:
        """Supprime les commentaires C et les litteraux de chaines,
        en preservant la longueur du texte (remplacement par des
        espaces) pour ne pas decaler les indices."""
        def _blank(match_text: str) -> str:
            return re.sub(r'[^\n]', ' ', match_text)

        text = re.sub(r'/\*.*?\*/', lambda m: _blank(m.group(0)),
                       text, flags=re.DOTALL)
        text = re.sub(r'//[^\n]*', lambda m: _blank(m.group(0)), text)
        text = re.sub(r'"[^"\\]*(?:\\.[^"\\]*)*"',
                       lambda m: _blank(m.group(0)), text)
        text = re.sub(r"'(?:\\.|[^'\\])'",
                       lambda m: _blank(m.group(0)), text)
        return text

    @staticmethod
    def _find_matching_paren(text: str, start: int) -> int:
        """start doit pointer sur un '('. Retourne l'index du ')'
        correspondant, ou -1 si non trouve."""
        if start >= len(text) or text[start] != '(':
            return -1
        depth = 0
        for j in range(start, len(text)):
            if text[j] == '(':
                depth += 1
            elif text[j] == ')':
                depth -= 1
                if depth == 0:
                    return j
        return -1

    @staticmethod
    def _find_matching_brace(text: str, start: int) -> int:
        """start doit pointer sur un '{'. Retourne l'index du '}'
        correspondant, ou len(text) si non trouve."""
        depth = 0
        for j in range(start, len(text)):
            if text[j] == '{':
                depth += 1
            elif text[j] == '}':
                depth -= 1
                if depth == 0:
                    return j
        return len(text)

    # ------------------------------------------------------------------
    # Extraction des appels + conditions
    # ------------------------------------------------------------------

    def _extract_calls_with_conditions(self,
                                        body: str) -> List[Tuple[str, str]]:
        """
        Analyse le corps d'une fonction (deja "clean") et retourne les
        appels avec leur condition englobante reelle.

        Regles :
        - if imbriques : toutes les conditions combinees avec ' && '
            ex: if(A){ if(B){ func() }}  =>  condition = 'A && B'
        - else         : negation du if correspondant si pas de
                         condition externe, sinon conditions externes
            ex: if(C){ } else { func() }          => '!(C)'
            ex: if(A){ if(B){} else { func() } }  => 'A'  (pas 'A && !(B)')
        - while / for  : condition de boucle ajoutee au contexte
        - switch/case  : label de case
        """
        clean = body
        results = []
        seen = set()

        i = 0
        n = len(clean)
        brace_depth = 0

        # Chaque entree : {'cond': str, 'depth': int, 'is_else': bool}
        cond_stack: List[dict] = []
        switch_stack: List[Tuple[str, int]] = []
        case_labels: dict = {}
        # Derniere condition 'if' vue a chaque profondeur de bloc,
        # conservee pour le 'else' eventuel.
        # Cle = brace_depth du CORPS du if (= profondeur_externe + 1)
        pending_if_cond: dict = {}

        while i < n:
            ch = clean[i]

            # ---- Accolades -----------------------------------------------
            if ch == '{':
                brace_depth += 1
                i += 1
                continue

            if ch == '}':
                # Retire les scopes appartenant a ce niveau
                cond_stack   = [e for e in cond_stack
                                if e['depth'] != brace_depth]
                switch_stack = [(s, d) for s, d in switch_stack
                                if d != brace_depth]
                case_labels.pop(brace_depth, None)
                # NE PAS effacer pending_if_cond : un 'else' peut suivre
                brace_depth -= 1
                i += 1
                continue

            # ---- switch --------------------------------------------------
            if ch == 's':
                sw_m = re.match(r'\bswitch\b', clean[i:])
                if sw_m and (i == 0 or not (clean[i-1].isalnum()
                                             or clean[i-1] == '_')):
                    kw_end    = i + len(sw_m.group(0))
                    paren_pos = clean.find('(', kw_end)
                    if paren_pos != -1 and paren_pos - kw_end <= 10:
                        close_pos = self._find_matching_paren(clean, paren_pos)
                        if close_pos != -1:
                            switch_expr = clean[paren_pos:close_pos + 1].strip()
                            switch_stack.append((switch_expr, brace_depth + 1))
                            i = close_pos + 1
                            continue

            # ---- case ----------------------------------------------------
            if ch == 'c':
                case_m = re.match(r'\bcase\b([^:]+):', clean[i:])
                if case_m and (i == 0 or not (clean[i-1].isalnum()
                                               or clean[i-1] == '_')):
                    case_val  = case_m.group(1).strip()
                    active_sw = next((s for s, d in reversed(switch_stack)
                                      if d == brace_depth), None)
                    label = (f"switch {active_sw} / case {case_val}"
                             if active_sw else f"case {case_val}")
                    case_labels[brace_depth] = label
                    i += len(case_m.group(0))
                    continue

            # ---- default -------------------------------------------------
            if ch == 'd':
                def_m = re.match(r'\bdefault\s*:', clean[i:])
                if def_m and (i == 0 or not (clean[i-1].isalnum()
                                              or clean[i-1] == '_')):
                    active_sw = next((s for s, d in reversed(switch_stack)
                                      if d == brace_depth), None)
                    label = (f"switch {active_sw} / default"
                             if active_sw else "default")
                    case_labels[brace_depth] = label
                    i += len(def_m.group(0))
                    continue

            # ---- else ----------------------------------------------------
            if ch == 'e':
                else_m = re.match(r'\belse\b', clean[i:])
                if else_m and (i == 0 or not (clean[i-1].isalnum()
                                               or clean[i-1] == '_')):
                    kw_end = i + len(else_m.group(0))
                    k = kw_end
                    while k < n and clean[k] in ' \t\r\n':
                        k += 1
                    if k < n and clean[k] == '{':
                        # else { ... } simple
                        if_cond = pending_if_cond.get(brace_depth + 1)
                        if if_cond is not None:
                            else_cond = f"!({if_cond})"
                            # Retire l'entree if de ce niveau si encore presente
                            cond_stack = [
                                e for e in cond_stack
                                if not (e['depth'] == brace_depth + 1
                                        and not e['is_else'])
                            ]
                            cond_stack.append({
                                'cond': else_cond,
                                'depth': brace_depth + 1,
                                'is_else': True
                            })
                            del pending_if_cond[brace_depth + 1]
                        # Pointer sur '{' pour la prochaine iteration
                        i = k
                        continue
                    else:
                        # else if(...) — sauter 'else', laisser 'if' traiter
                        i = k
                        continue

            # ---- if / while / for ----------------------------------------
            if ch in ('i', 'w', 'f'):
                kw_m = re.match(r'\b(if|while|for)\b', clean[i:])
                if kw_m and (i == 0 or not (clean[i-1].isalnum()
                                             or clean[i-1] == '_')):
                    kw        = kw_m.group(1)
                    kw_end    = i + len(kw_m.group(0))
                    paren_pos = clean.find('(', kw_end)
                    if paren_pos != -1 and paren_pos - kw_end <= 10:
                        close_pos = self._find_matching_paren(clean, paren_pos)
                        if close_pos != -1:
                            # Ignorer while(...); (fin d'un do { } while)
                            next_k = close_pos + 1
                            while next_k < n and clean[next_k] in ' \t\r\n':
                                next_k += 1
                            if (kw == 'while'
                                    and next_k < n and clean[next_k] == ';'):
                                i = next_k + 1
                                continue

                            # Expression interieure (sans les parentheses)
                            cond_expr = clean[paren_pos + 1:close_pos].strip()

                            if kw == 'if':
                                # Memoriser pour le else eventuel
                                pending_if_cond[brace_depth + 1] = cond_expr
                                cond_stack.append({
                                    'cond': cond_expr,
                                    'depth': brace_depth + 1,
                                    'is_else': False
                                })
                            else:
                                # while / for
                                cond_stack.append({
                                    'cond': f"{kw}({cond_expr})",
                                    'depth': brace_depth + 1,
                                    'is_else': False
                                })

                            i = close_pos + 1
                            continue

            # ---- Appels de fonctions -------------------------------------
            if ch.isalpha() or ch == '_':
                call_m = re.match(
                    r'([a-zA-Z_][a-zA-Z0-9_]*'
                    r'(?:::[a-zA-Z_][a-zA-Z0-9_]*)*)\s*\(',
                    clean[i:])
                if call_m and (i == 0 or not (
                        clean[i-1].isalnum() or clean[i-1] in ('_', ':'))):
                    call_name = call_m.group(1)
                    if call_name not in self._EXCLUDED_CALLS:
                        active = [e for e in cond_stack
                                  if e['depth'] <= brace_depth]
                        condition = self._build_condition_str(
                            active, case_labels, brace_depth)
                        key = (call_name, condition)
                        if key not in seen:
                            seen.add(key)
                            results.append(key)
                    i += len(call_m.group(0))
                    continue

            i += 1

        return results

    def _build_condition_str(self,
                              active_scopes: List[dict],
                              case_labels: dict,
                              brace_depth: int) -> str:
        """Construit la chaine de condition a partir des scopes actifs.

        Combine TOUTES les structures englobantes avec ' && ' :
          - if (cond)         -> 'cond'
          - else              -> '!(cond)' ou conditions externes
          - while (cond)      -> 'while(cond)'
          - for (init;c;inc)  -> 'for(init;c;inc)'
          - switch/case       -> 'switch(...)/case X'
          - imbrications      -> 'cond_A && while(B) && for(C) && case D'

        Regles pour else :
          * Conditions externes presentes -> uniquement ces conditions
            (la negation n'est pas ajoutee)
          * Pas de conditions externes    -> negation du if, ex '!(cond)'
        """
        # --- Partie 1 : conditions if / while / for ----------------------
        if not active_scopes:
            scope_cond = ''
        else:
            innermost = active_scopes[-1]
            if innermost['is_else']:
                else_depth  = innermost['depth']
                outer_conds = [e['cond'] for e in active_scopes
                               if e['depth'] < else_depth]
                if outer_conds:
                    # Seulement les conditions externes (sans la negation)
                    scope_cond = ' && '.join(outer_conds)
                else:
                    # Aucune condition externe -> negation du if
                    scope_cond = innermost['cond']
            else:
                # Combiner if + while + for avec &&
                scope_cond = ' && '.join(e['cond'] for e in active_scopes)

        # --- Partie 2 : label switch/case a la profondeur courante -------
        case_cond = case_labels.get(brace_depth, '')

        # --- Combinaison finale ------------------------------------------
        parts = [p for p in [scope_cond, case_cond] if p]
        return ' && '.join(parts) if parts else 'None'

    # ------------------------------------------------------------------
    # Informations CSC / CSU
    # ------------------------------------------------------------------

    def _file_info(self, file_path: Path) -> Dict[str, str]:
        """Retourne CSC et CSU a partir du fichier."""
        try:
            rel_parts = file_path.relative_to(self.source_dir).parts

            impl_index = None
            for idx, part in enumerate(rel_parts[:-1]):
                if part.lower() in self._IMPL_FOLDERS:
                    impl_index = idx
                    break

            if impl_index is not None and impl_index > 0:
                csc = rel_parts[impl_index - 1]
            elif len(rel_parts) >= 2:
                csc = rel_parts[0]
            else:
                csc = self.source_dir.name

        except ValueError:
            csc = file_path.parent.name

        stem = file_path.stem
        if stem.endswith('_code'):
            csu = stem[:-5]
        else:
            csu = stem

        return {'CSC': csc, 'CSU': csu}

    # ------------------------------------------------------------------
    # Generation des lignes Excel
    # ------------------------------------------------------------------

    def _to_excel_rows(self) -> List[List]:
        """Convertit les donnees extraites en lignes Excel"""
        rows = []

        for func_name in sorted(self.functions.keys()):
            info = self.functions[func_name]
            caller_info = self._file_info(info['file'])

            caller_csc = caller_info.get('CSC') or 'None'
            caller_csu = caller_info.get('CSU') or 'None'

            if not info['calls']:
                rows.append([
                    caller_csc,
                    caller_csu,
                    func_name,
                    'None',
                    'None',
                    'None',
                    'N/A (no calls in this function)'
                ])
                continue

            for callee_name, condition in info['calls']:
                callee_data = self.functions.get(callee_name)
                if callee_data:
                    callee_meta = self._file_info(callee_data['file'])
                    callee_csc = callee_meta.get('CSC') or 'None'
                    callee_csu = callee_meta.get('CSU') or 'None'
                else:
                    callee_csc = 'None'
                    callee_csu = 'None'

                cond_str = condition if condition is not None else 'None'

                rows.append([
                    caller_csc,
                    caller_csu,
                    func_name,
                    callee_csc,
                    callee_csu,
                    callee_name,
                    cond_str
                ])

        if not rows:
            rows.append(['None', 'None', 'None', 'None', 'None', 'None', 'None'])

        return rows


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        extractor = CCodeExtractor(sys.argv[1])
        result = extractor.extract()
        for row in result[:30]:
            print(row)