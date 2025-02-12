from vbaListener import vbaListener
from vbaParser import vbaParser


class VBA_L2(vbaListener):

    def __init__(self):
        # result

        self.names = []
        self.comments = []

        # state
        self.ids = {}
        self.variables = {}
        self.in_var_define = False
        self.in_type = False
        self.in_sub_define = False

    def finalize(self):
        # print("ids = " + ", ".join(self.ids.keys()))
        # print("variables = " + ", ".join(self.variables.keys()))

        for name in self.ids:
            if name not in self.variables:
                self.names.append(name)

    def exitAmbiguousIdentifier(self, ctx:vbaParser.AmbiguousIdentifierContext):
        name = ctx.getText()

        if self.in_var_define:
            if self.in_type:
                self.ids[name] = True
            else:
                self.variables[name] = True
            return

        if self.in_sub_define:
            if self.in_type:
                self.ids[name] = True
            else:
                self.variables[name] = True
            return

        self.ids[name] = True

    def enterVariableSubStmt(self, ctx:vbaParser.VariableSubStmtContext):
        self.in_var_define = True

    # Exit a parse tree produced by vbaParser#variableSubStmt.
    def exitVariableSubStmt(self, ctx:vbaParser.VariableSubStmtContext):
        self.in_var_define = False

    def exitComplexType(self, ctx:vbaParser.ComplexTypeContext):
        self.in_type = False

    def enterComplexType(self, ctx:vbaParser.ComplexTypeContext):
        self.in_type = True

    def enterEndOfStatement(self, ctx: vbaParser.EndOfStatementContext):
        self.in_sub_define = False

    def enterSubStmt(self, ctx: vbaParser.SubStmtContext):
        self.in_sub_define = True

    def enterFunctionStmt(self, ctx: vbaParser.FunctionStmtContext):
        self.in_sub_define = True

    def exitComment(self, ctx:vbaParser.CommentContext):
        line = ctx.getText()
        self.comments.append(line)
