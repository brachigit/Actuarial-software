from tree_sitter import Language, Parser

C_LANGUAGE = Language('build/my-languages.so', 'c')

parser = Parser()
parser.set_language(C_LANGUAGE)

from tree_sitter import Language, Parser

# יש להפעיל פעם אחת לבניית הספרייה
Language.build_library(
    'build/my-languages.so',
    ['tree-sitter-c']
)

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
            "if": "IF",
            "iferror": "IFERROR",
            "ifna": "IFNA",
            "xor": "XOR",
            "switch": "SWITCH",
            "true": "TRUE",
            "false": "FALSE",

            # Existing math
            "pow": "POWER",
            "sqrt": "SQRT",
            "abs": "ABS",
            "log": "LN",  # base e
            "exp": "EXP",
            "floor": "ROUNDDOWN",
            "ceil": "ROUNDUP",
            "max": "MAX",
            "min": "MIN",

            # Additional Excel math functions
            "log10": "LOG10",
            "mod": "MOD",
            "int": "INT",
            "round": "ROUND",
            "roundup": "ROUNDUP",
            "rounddown": "ROUNDDOWN",
            "trunc": "TRUNC",
            "sum": "SUM",
            "product": "PRODUCT",
            "average": "AVERAGE",
            "median": "MEDIAN",
            "even": "EVEN",
            "odd": "ODD",
            "rand": "RAND",
            "randbetween": "RANDBETWEEN",
            "pi": "PI",
            "sin": "SIN",
            "cos": "COS",
            "tan": "TAN",
            "asin": "ASIN",
            "acos": "ACOS",
            "atan": "ATAN",
            "atan2": "ATAN2",
            "degrees": "DEGREES",
            "radians": "RADIANS",
            "sign": "SIGN",
            "sqrtpi": "SQRTPI",
            "fact": "FACT",
            "combin": "COMBIN",
            "permut": "PERMUT",
            "gcd": "GCD",
            "lcm": "LCM",
            "quotient": "QUOTIENT",
            "power": "POWER",
            "ln": "LN",
            "log": "LOG",
            # Statistical functions
            "average": "AVERAGE",
            "mean": "AVERAGE",
            "median": "MEDIAN",
            "stdev": "STDEV.P",
            "stdevp": "STDEV.P",
            "stdevs": "STDEV.S",
            "var": "VAR.P",
            "vars": "VAR.S",
            "mode": "MODE.SNGL",
            "count": "COUNT",
            "counta": "COUNTA",
            "percentile": "PERCENTILE.INC",
            "rank": "RANK.EQ",
            "large": "LARGE",
            "small": "SMALL",
#data and time
            "now": "NOW",
            "today": "TODAY",
            "date": "DATE",
            "time": "TIME",
            "day": "DAY",
            "month": "MONTH",
            "year": "YEAR",
            "hour": "HOUR",
            "minute": "MINUTE",
            "second": "SECOND",
            "weekday": "WEEKDAY",
            "weeknum": "WEEKNUM",
            "edate": "EDATE",
            "eomonth": "EOMONTH",
            "datedif": "DATEDIF",  # not directly accessible in Excel UI, but supported
            "days": "DAYS",
            "networkdays": "NETWORKDAYS",
            "workday": "WORKDAY",
            "isoweeknum": "ISOWEEKNUM",
            "timevalue": "TIMEVALUE",
            "datevalue": "DATEVALUE",
            "yearfrac": "YEARFRAC",
#financial
            "pmt": "PMT",
            "fv": "FV",
            "pv": "PV",
            "rate": "RATE",
            "nper": "NPER",
            "npv": "NPV",
            "irr": "IRR",
            "xnpv": "XNPV",
            "xirr": "XIRR",
            "mirr": "MIRR",
            "ispmt": "ISPMT",
            "coupnum": "COUPNUM",
            "coupdaybs": "COUPDAYBS",
            "coupdays": "COUPDAYS",
            "coupdaysnc": "COUPDAYSNC",
            "price": "PRICE",
            "yield": "YIELD",
            "duration": "DURATION"

        }

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
            args = []
            for arg in args_node.named_children:
                args.append(self.convert_expression(arg, code))
            return f"{func_name}({', '.join(args)})"

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

    def convert(self, c_code: str) -> str:
        tree = self.parser.parse(bytes(c_code, "utf8"))
        root = tree.root_node
        func_node = root.children[0]
        body = func_node.child_by_field_name("body")
        code = c_code


        for child in body.children:
            print(f"Child type: {child.type}, text: {code[child.start_byte:child.end_byte]}")
            if child.type == 'if_statement':
                formula = self.convert_if_statement(child, code)
                self.excel_formula = formula
                return formula
            if child.type == 'return_statement':
                return self.convert_expression(child, code)

        return "NO_RETURN_FOUND"

    def convert_if_statement(self, node, code):
        condition_node = node.child_by_field_name("condition")
        consequence_node = node.child_by_field_name("consequence")
        alternative_node = node.child_by_field_name("alternative")

        condition = self.convert_expression(condition_node, code)
        if_true = self.extract_return_value(consequence_node, code)
        if_false = self.extract_return_value(alternative_node, code) if alternative_node else '""'

        return f"IF({condition}, {if_true}, {if_false})"

    def extract_return_value(self, node, code):
        if node.type == 'return_statement':
            if node.child_count > 1:
                expr_node = node.children[1]
                return self.get_expression_text(expr_node, code)
            else:
                print("לא נמצא ערך בתוך return")
                return ""

        elif node.type == 'compound_statement':
            for child in node.children:
                if child.type == 'return_statement':
                    return self.extract_return_value(child, code)
            print("לא נמצא return בתוך הבלוק")
            return ""

        elif node.type == 'else_clause':
            for child in node.children:
                if child.type == 'return_statement' or child.type == 'compound_statement':
                    return self.extract_return_value(child, code)
            print("לא נמצא return בתוך else_clause")
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


# ====================
# 🧪 Example Usage:
# ====================
if __name__ == "__main__":
    c_code = """
double calc() {
    if (life->retirement_age_lookup(1) > takeup_age)
        return 0.;
}

    """
    converter = CToExcelConverter()
    excel_formula = converter.convert(c_code)
    print("Excel Formula:", excel_formula)
