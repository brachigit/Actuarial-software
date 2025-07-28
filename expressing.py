from tree_sitter import Language, Parser

C_LANGUAGE = Language('build/my-languages.so', 'c')

parser = Parser()
parser.set_language(C_LANGUAGE)


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
        }

    def convert(self, c_code: str) -> str:
        tree = self.parser.parse(bytes(c_code, "utf8"))
        root = tree.root_node
        func_node = root.children[0]
        body = func_node.child_by_field_name("body")
        code = c_code

        # נבנה רשימה של if-statements וה-return שבסוף
        statements = []
        for child in body.named_children:
            if child.type == 'if_statement':
                statements.append(child)
            elif child.type == 'return_statement':
                statements.append(child)

        formula = self.build_nested_if(statements, code)
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
                return f"({left} = {right})"
            elif op == '!=':
                return f"({left} <> {right})"
            elif op == '&&':
                return f"AND({left}, {right})"
            elif op == '||':
                conditions = self.flatten_or_conditions(node, code)
                return self.build_or_expression(conditions)

            else:
                return f"({left} {op} {right})"


        elif kind == 'call_expression':
            function_node = children[0]
            args_node = children[1]
            func_name = self.convert_expression(function_node, code)

            # המרה של שם הפונקציה בהתאם למפה, אם לא קיים, נשאיר כמו שהוא
            excel_func_name = self.func_map.get(func_name.lower(), func_name)

            args = []
            for arg in args_node.named_children:
                args.append(self.convert_expression(arg, code))
            return f"{excel_func_name}({', '.join(args)})"


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
            return f"({self.convert_expression(children[1], code)})"

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

        return code[node.start_byte:node.end_byte]

    def convert_if_statement(self, node, code):
        condition = self.convert_expression(node.child_by_field_name('condition'), code)
        if_true = node.child_by_field_name('consequence')
        if_false = node.child_by_field_name('alternative')

        # קבל את הערך של ה-if_true
        true_result = self.extract_return_value(if_true, code)
        if true_result == '""':
            assign = self.find_assignment_in_node(if_true)
            if assign:
                true_result = self.convert_expression(assign.children[1], code)

        # קבל את הערך של ה-if_false
        false_result = self.extract_return_value(if_false, code) if if_false else '""'
        if false_result == '""' and if_false:
            assign = self.find_assignment_in_node(if_false)

            if assign:
                false_result = self.convert_expression(assign.children[1], code)

        return f'IF(({condition}), {true_result}, {false_result})'

    def build_nested_if(self, statements, code):
        """
        מקבל רשימת תנאים ויוצר נוסחת IF מקוננת אחת.
        """
        default_result = '""'

        # עוברים אחורה כדי לבנות קינון פנימי -> חיצוני
        for node in reversed(statements):
            if node.type == 'return_statement':
                default_result = self.convert_expression(node, code)
            elif node.type == 'if_statement':
                condition = self.convert_expression(node.child_by_field_name('condition'), code)
                if_true = self.extract_return_value(node.child_by_field_name('consequence'), code)

                # יש גם else עם return?
                if node.child_by_field_name('alternative') is not None:
                    if_false = self.extract_return_value(node.child_by_field_name('alternative'), code)
                else:
                    if_false = default_result

                default_result = f'IF(({condition}), {if_true}, {if_false})'

        return default_result

    def extract_return_value(self, node, code):
        if node.type == 'return_statement':
            if node.child_count > 1:
                expr_node = node.children[1]
                return self.get_expression_text(expr_node, code)
            else:
                return ""

        elif node.type == 'compound_statement':  # מתאר בלוק פקודות
            for child in node.children:
                if child.type == 'return_statement':
                    return self.extract_return_value(child, code)
                elif child.type == 'if_statement':
                    nested_result = self.convert_if_statement(child, code)
                    if nested_result:
                        return nested_result
                elif child.type == 'expression_statement':
                    # אם זה ביטוי השמה נשלוף רק את צד ימין
                    if len(child.children) > 0 and child.children[0].type == 'assignment_expression':
                        assign_expr = self.convert_expression(child.children[0].children[2], code)
                        if assign_expr:
                            return assign_expr

            return ""



        elif node.type == 'else_clause':

            for child in node.children:

                if child.type == 'return_statement' or child.type == 'compound_statement':
                    return self.extract_return_value(child, code)


                elif child.type == 'if_statement':
                    nested_result = self.convert_if_statement(child, code)
                    if nested_result:
                        return nested_result
                elif child.type == 'expression_statement':
                    # אם זה ביטוי השמה נשלוף רק את צד ימין
                    if len(child.children) > 0 and child.children[0].type == 'assignment_expression':
                        assign_expr = self.convert_expression(child.children[0].children[2], code)
                        if assign_expr:
                            return assign_expr
            return ""

        else:
            print(f"צומת אינו return_statement או block: {node.type}")
            return ""

        # 🔵 פונקציית עזר לדיבוג: הדפסת שדות אפשריים

    def print_node_fields(self, node, code):
        field_names = [
            'condition', 'consequence', 'alternative', 'value',
            'declarator', 'body', 'argument', 'arguments'
        ]
        for name in field_names:
            child = node.child_by_field_name(name)
            if child is not None:
                print(f"Field: {name}, Type: {child.type}, Text: {code[child.start_byte:child.end_byte]}")

    def get_expression_text(self, node, code):
        return code[node.start_byte:node.end_byte]

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

        # אם זו השמה - מחזירים מיד
        if node.type == 'assignment_expression':
            return node

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

        # אם זו else או if או כל דבר אחר – נכנסים רקורסיבית
        for child in node.children:
            result = self.find_assignment_in_node(child)
            if result:
                return result

        return None


# ====================
# 🧪 Example Usage:
# ====================
if __name__ == "__main__":
    c_code = """
double calc() {
   if (t <= commence_period_w || t >  maturity_period_ann)
 return 0.0;
 if (annuity_pmt_curr_tot == 0)
 return 0.0;
 return - pmt_total(t) - expense_ren_perc_post_ret(t);
 
 }
    """

    converter = CToExcelConverter()
    excel_formula = converter.convert(c_code)
    print("Excel Formula:", excel_formula)
