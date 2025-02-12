from vbaListener import vbaListener
from vbaParser import vbaParser


class VBA_L1(vbaListener):

    def __init__(self):
        # result
        self.subs = []
        self.ids = []
        self.namespace = ""
        self.global_code = []

        # state
        self.in_attr = False
        self.in_name_attr = False
        self.in_global = False

    def exitSubStmt(self, ctx: vbaParser.SubStmtContext):
        content = ctx.getText()
        self.subs.append(content)
        ident = ctx.ambiguousIdentifier()
        name = ident.getText()
        self.ids.append(name)

    def exitFunctionStmt(self, ctx: vbaParser.FunctionStmtContext):
        content = ctx.getText()
        self.subs.append(content)
        ident = ctx.ambiguousIdentifier()
        name = ident.getText()
        self.ids.append(name)

    def enterAttributeStmt(self, ctx:vbaParser.AttributeStmtContext):
        self.in_attr = True
        self.in_name_attr = False

    def exitAttributeStmt(self, ctx:vbaParser.AttributeStmtContext):
        self.in_attr = False
        self.in_name_attr = False

    def exitImplicitCallStmt_InStmt(self, ctx:vbaParser.ImplicitCallStmt_InStmtContext):
        if self.in_attr:
            name = ctx.getText()
            if name.lower() == "vb_name":
                self.in_name_attr = True

    def exitLiteral(self, ctx: vbaParser.LiteralContext):
        if self.in_name_attr:
            name = ctx.getText().strip('\"')
            self.ids.append(name)
            self.namespace = name

    def enterModuleDeclarationsElement(self, ctx:vbaParser.ModuleDeclarationsElementContext):
        self.in_global = True

    # Exit a parse tree produced by vbaParser#moduleDeclarationsElement.
    def exitModuleDeclarationsElement(self, ctx:vbaParser.ModuleDeclarationsElementContext):
        content = ctx.getText()
        self.global_code.append(content)
        self.in_global = False

    def exitAmbiguousIdentifier(self, ctx:vbaParser.AmbiguousIdentifierContext):
        if self.in_global:
            name = ctx.getText()
            self.ids.append(name)
