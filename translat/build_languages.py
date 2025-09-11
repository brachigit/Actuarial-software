from tree_sitter import Language  # type: ignore
import os

# מיקום לקובץ המשותף שייבנה
so_path = 'build/my-languages.so'

# הנתיב לקוד המקור של tree-sitter-c וה־cpp
grammar_paths = [
    'tree-sitter-c',    # צריך לשכפל את זה (ראה שלב הבא)
    'tree-sitter-cpp'   # גם את זה
]

# בונה את הספרייה
Language.build_library(
    # פלט
    so_path,
    # קלט
    grammar_paths
)
