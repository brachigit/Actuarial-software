from tree_sitter import Language, Parser  # type: ignore
from formulas import is_valid_excel_formula
from write_to_excel import  write_to_excel
import os

C_LANGUAGE = Language('build/my-languages.so', 'c')

parser = Parser()
parser.set_language(C_LANGUAGE)


def wrap(expr: str) -> str:
    expr = expr.strip()

    # אם כבר יש סוגריים מסביב לכל הביטוי - נוריד אותן
    while expr.startswith("(") and expr.endswith(")"):
        # נבדוק אם הסוגריים באמת עוטפות את כל הביטוי
        depth = 0
        is_outer = True
        for i, ch in enumerate(expr):
            if ch == '(':
                depth += 1
            elif ch == ')':
                depth -= 1
                if depth == 0 and i != len(expr) - 1:
                    is_outer = False
                    break
        if is_outer:
            expr = expr[1:-1].strip()
        else:
            break

    # לא נוסיף סוגריים אם זה ביטוי פשוט
    if any(op in expr for op in [' ', '+', '-', '*', '/', '<', '>', '=', ',', 'AND', 'OR']):
        return expr
    return expr


class CToExcelConverter:
    def __init__(self):
        self.parser = parser

        self.func_map = {
            "logical_and": "AND",
            "logical_or": "OR",
            "logical_not": "NOT",
            "plus": "+",
            "minus": "-",
            "star": "*",
            "slash": "/",
            "percent": "MOD",
            "greater": ">",
            "greater_equal": ">=",
            "less": "<",
            "less_equal": "<=",
            "equal_equal": "=",
            "not_equal": "<>",
            "and": "AND",
            "or": "OR",
            "not": "NOT",
            "negate": "-",
            "max": "MAX",
            "min": "MIN",
            "abs": "ABS",
            "pow": "POWER",
            "sqrt": "SQRT",
            "exp": "EXP",
            "log": "LN",
            "log10": "LOG10",
            "floor": "ROUNDDOWN",
            "ceil": "ROUNDUP",
            "round": "ROUND",
            "if": "IF",
            "sum": "SUM",
            "average": "AVERAGE",
            "count": "COUNT",
            "median": "MEDIAN",
            "mode": "MODE",
            "stdev": "STDEV",
            "var": "VAR",
            "large": "LARGE",
            "small": "SMALL",
            "left": "LEFT",
            "right": "RIGHT",
            "mid": "MID",
            "len": "LEN",
            "find": "FIND",
            "search": "SEARCH",
            "substitute": "SUBSTITUTE",
            "concat": "CONCAT",
            "upper": "UPPER",
            "lower": "LOWER",
            "trim": "TRIM",
            "now": "NOW",
            "today": "TODAY",
            "day": "DAY",
            "month": "MONTH",
            "year": "YEAR",
            "hour": "HOUR",
            "minute": "MINUTE",
            "second": "SECOND",
            "text": "TEXT",
            "eq": "="
        }
        self.variables = {}

    def convert(self, c_code: str) -> str:
        # אתחול מחדש של מילון המשתנים
        self.variables = {}

        tree = self.parser.parse(bytes(c_code, "utf8"))
        root = tree.root_node
        func_node = root.children[0]
        body = func_node.child_by_field_name("body")
        self.code = c_code

        tree = self.parser.parse(bytes(self.code, "utf8"))
        root_node = tree.root_node
        # נבנה רשימה של if-statements וה-return שבסוף
        # נבנה רשימה של כל ה־statements לפי הסדר האמיתי
        statements = []
        for child in body.named_children:
            if child.type in ('if_statement', 'return_statement', 'declaration', 'expression_statement'):
                statements.append(child)
        formula = self.build_nested_if(statements, self.code)

        self.excel_formula = formula
        return formula

    def convert_expression(self, node, code):
        kind = node.type
        children = node.children

        if kind == 'binary_expression':
            left = self.convert_expression(children[0], code)
            op_node = children[1]
            right = self.convert_expression(children[2], code)
            op = code[op_node.start_byte:op_node.end_byte]

            # ניהול אופרטורים לוגיים והשוואות
            if op == '==':
                return wrap(f"{left} = {right}")
            elif op == '!=':
                return wrap(f"{left} <> {right}")
            elif op == '&&':
                return wrap(f"AND{left}, {right}")
            elif op == '||':
                conditions = self.flatten_or_conditions(node, code)
                return self.build_or_expression(conditions)
            else:
                return wrap(f"{left} {op} {right}")
        elif kind == 'call_expression':
            function_node = children[0]
            args_node = children[1]
            func_name = self.convert_expression(function_node, code)
            args = []

            for arg in args_node.named_children:
                args.append(self.convert_expression(arg, code))

            # ✅ טיפול מיוחד בפונקציה eq(a,b)
            if func_name.lower() == "eq":
                return f"({args[0]} = {args[1]})"

            # המרה של שם הפונקציה בהתאם למפה
            excel_func_name = self.func_map.get(func_name.lower(), func_name)
            return f"{excel_func_name}({', '.join(args)})"

        elif kind == 'assignment_expression':

            left = children[0]
            right = children[2]
            var_name = code[left.start_byte:left.end_byte]
            value_formula = self.convert_expression(right, code)
            self.handle_assignment(var_name, value_formula)
            return ''  # לא מחזירים כלום כאן - משתנה נשמר

        elif kind == 'field_expression':
            # example: numbers->age
            # left is the file (numbers), right is the field (age)
            left = self.convert_expression(children[0], code)
            right = code[children[2].start_byte:children[2].end_byte]
            return f'{left}!{right}'

        elif kind == 'conditional_expression':
            condition = self.convert_expression(children[0], code)
            true_expr = self.convert_expression(children[2], code)
            false_expr = self.convert_expression(children[4], code)
            return f"IF({condition}, {true_expr}, {false_expr})"

        elif kind == 'identifier' or kind == 'number_literal':
            return code[node.start_byte:node.end_byte]

        elif kind == 'parenthesized_expression':
            return wrap(f"{self.convert_expression(children[1], code)}")

        elif kind == 'return_statement':
            return self.convert_expression(children[1], code)

        elif kind == 'unary_expression':
            operator_node = children[0]
            operand_node = children[1]
            operand = self.convert_expression(operand_node, code)
            operator = code[operator_node.start_byte:operator_node.end_byte]

            if operator == '!':
                return f"NOT({operand})"
            elif operator == '-':
                return f"-({operand})"
            elif operator == '+':
                return f"+({operand})"
            else:
                return f"{operator}({operand})"

        elif kind == 'subscript_expression':
            # דוגמה: life->free_inv_prop_t[inv_year]
            array_node = children[0]
            index_node = children[2]

            # אם זה access דרך ->
            if array_node.type == 'field_expression':
                left = self.convert_expression(array_node.children[0], code)  # life
                right = code[array_node.children[2].start_byte:array_node.children[2].end_byte]  # free_inv_prop_t
                array_ref = f"{left}!{right}"
            else:
                # גישה רגילה למערך
                array_ref = self.convert_expression(array_node, code)

            index_ref = self.convert_expression(index_node, code)
            # באקסל צריך אינדקס החל מ-1, ב־C הוא מ-0
            return f"INDEX({array_ref}, {index_ref} + 1)"
        return code[node.start_byte:node.end_byte]

    def convert_if_statement(self, node):
        condition_node = node.child_by_field_name("condition")
        consequence_node = node.child_by_field_name("consequence")
        alternative_node = node.child_by_field_name("alternative")

        condition = self.convert_expression(condition_node, self.code)

        def convert_block(block_node):
            expressions = []
            if block_node.type == "compound_statement":
                for child in block_node.children:
                    if child.type in ["{", "}"]:
                        continue
                    expr = self.convert_statement(child)
                    if expr:
                        expressions.append(expr)
            else:
                expr = self.convert_statement(block_node)
                if expr:
                    expressions.append(expr)
            return ", ".join(expressions) if expressions else ""

        consequence = convert_block(consequence_node)
        alternative = convert_block(alternative_node) if alternative_node else None

        if alternative:
            return f"IF({condition}, {consequence}, {alternative})"
        else:
            return f"IF({condition}, {consequence})"

    def convert_statement(self, node):
        if node.type == "if_statement":
            result = self.convert_if_statement(node)

            # 🔹 טיפול ב־alternative (else) כדי לעדכן השמות
            alternative_node = node.child_by_field_name("alternative")
            if alternative_node:
                self.extract_assignments(alternative_node, self.code)
            return result

        elif node.type == "return_statement":
            return self.extract_return_value(node, self.code)

        elif node.type == "expression_statement":
            expr_node = node.child_by_field_name("expression")
            if expr_node is not None:
                if expr_node.type == "assignment_expression":
                    self.convert_expression(expr_node, self.code)
                    return ""  # משתנה נשמר, אין נוסחה להחזיר
                else:
                    return self.convert_expression(expr_node, self.code)
        elif node.type == "assignment_expression":
            self.convert_expression(node, self.code)
            return ""
        elif node.type == "compound_statement":
            results = []
            for child in node.named_children:
                result = self.convert_statement(child)
                if result:
                    results.append(result)
            return ", ".join(results)

        elif node.type == 'declaration':
            declarator = node.child_by_field_name("declarator")
            value_node = node.child_by_field_name("value")
            # טיפול במבנה init_declarator אם יש
            if declarator and declarator.type == "init_declarator":
                var_id = declarator.child_by_field_name("declarator")
                value_node = declarator.child_by_field_name("value")
                if var_id and value_node:
                    var_name = self.get_expression_text(var_id, self.code)
                    value = self.convert_expression(value_node, self.code)
                    self.handle_assignment(var_name, value)
                return

            # טיפול רגיל בהצהרה עם ערך
            if declarator and value_node:
                var_name = self.get_expression_text(declarator, self.code)
                # כאן חשוב מאוד לשלוח את כל value_node ל-convert_expression
                value = self.convert_expression(value_node, self.code)
                self.handle_assignment(var_name, value)
            # טיפול רגיל
        else:
            # כל דבר אחר – ננסה לטפל כביטוי
            return self.convert_expression(node, self.code)

    def handle_assignment(self, name, value):
        self.variables[name] = value
        return None  # לא מחזיר מיידית נוסחה אלא מחכה לבניית LET בסוף

    def extract_return_value(self, node, code):
        if node.type == 'return_statement':
            if node.child_count > 1:
                expr_node = node.children[1]
                value_text = self.get_expression_text(expr_node, code)

                return self.convert_expression(expr_node, code)



            else:
                return ""

        elif node.type == 'compound_statement':
            for child in node.children:
                if child.type == 'return_statement':
                    return self.extract_return_value(child, code)
                elif child.type == 'if_statement':
                    nested_result = self.convert_if_statement(child)
                    if nested_result:
                        return nested_result
                elif child.type == 'expression_statement':
                    if len(child.children) > 0 and child.children[0].type == 'assignment_expression':
                        self.convert_expression(child.children[0], code)
                        return ''
            return ""

        elif node.type == 'else_clause':
            for child in node.children:
                if child.type in ('return_statement', 'compound_statement'):
                    return self.extract_return_value(child, code)
                elif child.type == 'if_statement':
                    nested_result = self.convert_if_statement(child, code)
                    if nested_result:
                        return nested_result
                elif child.type == 'expression_statement':
                    if len(child.children) > 0 and child.children[0].type == 'assignment_expression':
                        assign_expr = self.convert_expression(child.children[0].children[2], code)
                        if assign_expr:
                            return assign_expr
            return ""
        else:
            return ""

        # 🔵 פונקציית עזר לדיבוג: הדפסת שדות אפשריים


    def get_expression_text(self, node, code):
        if node.type == "identifier":
            return code[node.start_byte:node.end_byte]
        return code[node.start_byte:node.end_byte]

    def extract_statements(self, node):
        """
        מקבלת node של consequence / alternative ומחזירה רשימת statements
        """
        stmts = []
        if node is None:
            return stmts
        if node.type == 'compound_statement':
            for child in node.named_children:
                if child.type in ('if_statement', 'return_statement', 'declaration', 'expression_statement'):
                    stmts.append(child)
        else:
            if node.type in ('if_statement', 'return_statement', 'declaration', 'expression_statement'):
                stmts.append(node)
        return stmts

    def merge_variables(self, vars_before, vars_true, vars_false, condition):
        merged = {}
        all_vars = set(vars_before.keys()) | set(vars_true.keys()) | set(vars_false.keys())

        for var in all_vars:
            val_before = vars_before.get(var, '""')
            val_true = vars_true.get(var, val_before)
            val_false = vars_false.get(var, val_before)

            if val_true == val_false:
                merged[var] = val_true
            else:
                merged[var] = f"IF({condition}, {val_true}, {val_false})"
        return merged

    def build_nested_if(self, statements, code):
        """
        בונה IF מקוננים מהצהרות ומשמרת את כל ההשמות ב-LET יחיד בסוף.
        """
        default_result = '""'
        variables_before = dict(self.variables)  # מצב משתנים לפני העיבוד

        for node in reversed(statements):
            if node.type == 'return_statement':
                # קביעת ערך ברירת מחדל לחזרה
                default_result = self.convert_expression(node, code)

            elif node.type == 'if_statement':
                condition_node = node.child_by_field_name('condition')
                consequence_node = node.child_by_field_name('consequence')
                alternative_node = node.child_by_field_name('alternative')

                condition = self.convert_expression(condition_node, code)

                # שמירת מצב משתנים לפני הענפים
                vars_before_if = dict(self.variables)

                # ===== TRUE branch =====
                self.variables = dict(vars_before_if)
                consequence_statements = self.extract_statements(consequence_node)
                consequence_formula = self.build_nested_if(consequence_statements, code)
                vars_true = dict(self.variables)

                # ===== FALSE branch =====
                self.variables = dict(vars_before_if)
                if alternative_node:
                    alternative_statements = self.extract_statements(alternative_node)
                    alternative_formula = self.build_nested_if(alternative_statements, code)
                    vars_false = dict(self.variables)
                else:
                    alternative_formula = default_result
                    vars_false = dict(self.variables)

                # מיזוג ערכי משתנים לענף יחיד
                merged_vars = self.merge_variables(vars_before_if, vars_true, vars_false, condition)
                self.variables = merged_vars

                # עדכון הפלט
                default_result = f"IF({condition}, {consequence_formula}, {alternative_formula})"

            elif node.type == 'expression_statement':
                inner = node.children[0] if node.children else None
                if inner and inner.type == 'assignment_expression':
                    self.convert_expression(inner, code)

            elif node.type == 'declaration':
                self.extract_assignments(node, code)

            elif node.type == 'compound_statement':
                self.extract_assignments(node, code)

        # יצירת LET יחיד בסוף
        if self.variables:
            bindings = ', '.join(f'{k}, {v}' for k, v in self.variables.items())
            return f'LET({bindings}, {default_result})'
        else:
            return default_result

    def build_or_expression(self, conditions: list[str]) -> str:
        if not conditions:
            return ""
        if len(conditions) == 1:
            return conditions[0]
        return f'OR({", ".join(conditions)})'

    def flatten_or_conditions(self, node, code):
        """
        מקבלת עץ של תנאי '||' ומחזירה רשימת ביטויים שטוחים.
        """
        if node.type == 'binary_expression':
            children = node.children
            op_node = children[1]
            op = code[op_node.start_byte:op_node.end_byte]

            if op == '||':
                left_conditions = self.flatten_or_conditions(children[0], code)
                right_conditions = self.flatten_or_conditions(children[2], code)
                return left_conditions + right_conditions

        # אם זה לא || - נחזיר את הביטוי כמו שהוא
        return [self.convert_expression(node, code)]

    def find_assignment_in_node(self, node):

        if node is None:
            return None

        # אם זו השמה - מחזירים את שם המשתנה ואת הצומת
        if node.type == 'assignment_expression':
            left_node = node.child_by_field_name('left')
            if left_node is not None and left_node.type == 'identifier':
                var_name = left_node.text.decode('utf-8')
                return (var_name, node)
            else:
                return (None, node)  # למקרה חריג שאין left ברור

        # אם זו שורת קוד שמכילה השמה
        if node.type == 'expression_statement':
            for child in node.children:
                result = self.find_assignment_in_node(child)
                if result:
                    return result

        # אם זו בלוק של קוד
        if node.type == 'compound_statement':
            for child in node.children:
                result = self.find_assignment_in_node(child)
                if result:
                    return result

        if node.type == "declaration":
            declarator = node.child_by_field_name("declarator")
            value_node = node.child_by_field_name("value")
            if declarator and value_node:
                var_name = self.get_expression_text(declarator, self.code)
                return var_name, value_node

        # אם זו else או if או כל דבר אחר – נכנסים רקורסיבית
        for child in node.children:
            result = self.find_assignment_in_node(child)
            if result:
                return result

        return None

    def extract_assignments(self, block_node, code):
        if block_node.type == 'compound_statement':
            for child in block_node.children:
                self.extract_assignments(child, code)  # רקורסיה פנימה

        elif block_node.type == 'expression_statement' and len(block_node.children) > 0:
            inner = block_node.children[0]
            if inner.type == 'assignment_expression':
                self.convert_expression(inner, code)

        elif block_node.type == 'declaration':
            declarator = block_node.child_by_field_name("declarator")
            value_node = block_node.child_by_field_name("value")

            if declarator and declarator.type == "init_declarator":
                var_id = declarator.child_by_field_name("declarator")
                value_node = declarator.child_by_field_name("value")
                if var_id and value_node:
                    var_name = self.get_expression_text(var_id, code)
                    value = self.convert_expression(value_node, code)
                    self.handle_assignment(var_name, value)
                return

            if declarator and value_node:
                var_name = self.get_expression_text(declarator, code)
                value = self.convert_expression(value_node, code)
                self.handle_assignment(var_name, value)



# ====================
# 🧪 Example Usage:
# ====================
if __name__ == "__main__":
    c_code = """
double rectangle_area() 
{  if (t <= commence_period_w || t >  maturity_period_ann)
 return 0.0;
 if (annuity_pmt_curr_tot == 0)
 return 0.0;
 return - pmt_total(t) - expense_ren_perc_post_ret(t); }

    """

    converter = CToExcelConverter()
    excel_formula = converter.convert(c_code)
    print( is_valid_excel_formula(excel_formula))
    print("Excel Formula:", excel_formula)

    excel_formula='=IF(H1 >, "Yes", "No")'
    excel_file = "AuditReport_With_Lookups_extracted_20250827_032233.xlsx"
    excel_path = os.path.join("outputs", excel_file)
    print("Excel path:", excel_path)
    print("File exists:", os.path.exists(excel_path))
    write_to_excel(excel_formula,"D2","AuditReport_With_Lookups_extracted_20250827_032233.xlsx","Data")