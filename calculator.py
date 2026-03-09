import ast
import operator
import re

# Maps AST operator nodes to their corresponding Python operator functions.
# Only arithmetic operators are supported; anything else raises an error.
OPERATORS = {
    ast.Add: operator.add,
    ast.Sub: operator.sub,
    ast.Mult: operator.mul,
    ast.Div: operator.truediv,
    ast.Pow: operator.pow,
    ast.USub: operator.neg,  # unary minus, e.g. -5
}

def insert_implicit_mul(expression):
    """Rewrite implicit multiplication so the AST parser can handle it.

    Handles three cases:
      - number followed by '(':   3(x)  → 3*(x)
      - ')' followed by '(':      )(    → )*(
      - ')' followed by number:   )3    → )*3
    """
    # number immediately before '(' → insert '*'
    expression = re.sub(r'(\d+\.?\d*)\s*\(', r'\1*(', expression)
    # closing paren immediately before opening paren → insert '*'
    expression = re.sub(r'\)\s*\(', r')*(', expression)
    # closing paren immediately before a number → insert '*'
    expression = re.sub(r'\)\s*(\d+\.?\d*)', r')*\1', expression)
    return expression

def eval_expr(node):
    """Recursively evaluate an AST node, returning a numeric result.

    Only allows numeric constants and the operators defined in OPERATORS.
    Using ast.parse instead of eval() prevents arbitrary code execution.
    """
    if isinstance(node, ast.Constant):
        # Leaf node: a literal number
        return node.value
    if isinstance(node, ast.BinOp):
        # Binary operation: left OP right (e.g. 3 + 4)
        op = OPERATORS.get(type(node.op))
        if op is None:
            raise ValueError("Unsupported operator")
        left = eval_expr(node.left)
        right = eval_expr(node.right)
        # Catch division by zero before calling the operator
        if isinstance(node.op, ast.Div) and right == 0:
            raise ZeroDivisionError
        return op(left, right)
    if isinstance(node, ast.UnaryOp):
        # Unary operation: OP operand (e.g. -5)
        op = OPERATORS.get(type(node.op))
        if op is None:
            raise ValueError("Unsupported operator")
        return op(eval_expr(node.operand))
    raise ValueError("Invalid expression")

def calculate(expression):
    """Preprocess and evaluate a math expression string, returning a number."""
    expression = insert_implicit_mul(expression)
    tree = ast.parse(expression, mode="eval")
    return eval_expr(tree.body)

def main():
    """Run the interactive calculator REPL."""
    print("Calculator — supports +, -, *, /, ** and parentheses")
    print("Type 'quit' to exit\n")

    while True:
        expression = input(">>> ").strip()

        # Allow several common exit commands
        if expression.lower() in ("quit", "exit", "q"):
            print("Goodbye!")
            break

        if not expression:
            continue

        try:
            result = calculate(expression)
            # Display whole-number floats without a decimal point (e.g. 4.0 → 4)
            print(f"= {int(result) if isinstance(result, float) and result == int(result) else result}\n")
        except ZeroDivisionError:
            print("Error: Division by zero\n")
        except Exception:
            print("Error: Invalid expression\n")

if __name__ == "__main__":
    main()
