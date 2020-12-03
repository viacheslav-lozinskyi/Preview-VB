
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System.Collections;
using System.IO;
using System.Linq;

namespace resource.preview
{
    internal class VSPreview : cartridge.AnyPreview
    {
        internal class HINT
        {
            public static string DATA_TYPE = "[[Data type]]";
            public static string METHOD_TYPE = "[[Method type]]";
        }

        protected override void _Execute(atom.Trace context, string url)
        {
            var a_Context = VisualBasicSyntaxTree.ParseText(File.ReadAllText(url)).WithFilePath(url).GetRoot();
            var a_IsFound = GetProperty(NAME.PROPERTY.DEBUGGING_SHOW_PRIVATE) != 0;
            if (a_Context == null)
            {
                return;
            }
            {
                context.
                    SetState(NAME.STATE.HEADER).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, 1, "[[Info]]");
                {
                    context.
                        SetValue(url).
                        Send(NAME.SOURCE.PREVIEW, NAME.TYPE.VARIABLE, 2, "[[File Name]]");
                    context.
                        SetValue(a_Context.GetText().Length.ToString()).
                        Send(NAME.SOURCE.PREVIEW, NAME.TYPE.VARIABLE, 2, "[[File Size]]");
                    context.
                        SetValue(a_Context.Language).
                        Send(NAME.SOURCE.PREVIEW, NAME.TYPE.VARIABLE, 2, "[[Language]]");
                }
            }
            if (a_Context.DescendantNodes().OfType<ImportsStatementSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<ImportsStatementSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, 1, "[[Dependencies]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<ImportsStatementSyntax>())
                {
                    __Execute(a_Context1, 2, context, url);
                }
            }
            if (a_Context.DescendantNodes().OfType<ClassBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<ClassBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, 1, "[[Classes]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<ClassBlockSyntax>())
                {
                    __Execute(a_Context1, 2, context, url, a_IsFound);
                }
            }
            if (a_Context.DescendantNodes().OfType<StructureBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<StructureBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, 1, "[[Structs]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<StructureBlockSyntax>())
                {
                    __Execute(a_Context1, 2, context, url, a_IsFound);
                }
            }
            if (a_Context.DescendantNodes().OfType<EnumBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<EnumBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, 1, "[[Enums]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<EnumBlockSyntax>())
                {
                    __Execute(a_Context1, 2, context, url, a_IsFound);
                }
            }
            if (a_Context.DescendantNodes().OfType<MethodBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<MethodBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FOLDER, 1, "[[Functions]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<MethodBlockSyntax>())
                {
                    __Execute(a_Context1, 2, context, url, true, a_IsFound);
                }
            }
            if (a_Context.GetDiagnostics().Any())
            {
                context.
                    SetState(NAME.STATE.FOOTER).
                    SetComment(__GetArraySize(a_Context.GetDiagnostics())).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.ERROR, 1, "[[Diagnostics]]");
                foreach (var a_Context1 in a_Context.GetDiagnostics())
                {
                    __Execute(a_Context1, 2, context, url);
                }
            }
            if (GetState() == STATE.CANCEL)
            {
                context.
                    SendWarning(1, NAME.WARNING.TERMINATED);
            }
        }

        private static void __Execute(Diagnostic node, int level, atom.Trace context, string url)
        {
            context.
                SetState(__GetSeverity(node)).
                SetUrlLine(__GetLine(node.Location)).
                SetUrlPosition(__GetPosition(node.Location)).
                SetUrl(url).
                SetLink("https://www.bing.com/search?q=" + node.Id).
                Send(NAME.SOURCE.PREVIEW, NAME.TYPE.INFO, level, node.Descriptor?.MessageFormat?.ToString());
        }

        private static void __Execute(ImportsStatementSyntax node, int level, atom.Trace context, string url)
        {
            context.
                SetComment("Imports").
                SetCommentHint(HINT.DATA_TYPE).
                SetUrlLine(__GetLine(node.GetLocation())).
                SetUrlPosition(__GetPosition(node.GetLocation())).
                SetUrl(url).
                Send(NAME.SOURCE.PREVIEW, NAME.TYPE.INFO, level, node.ImportsClauses.ToString());
        }

        private static void __Execute(ClassBlockSyntax node, int level, atom.Trace context, string url, bool isShowPrivate)
        {
            if (__IsEnabled(node.BlockStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment("Class").
                    SetCommentHint(HINT.DATA_TYPE).
                    SetUrlLine(__GetLine(node.GetLocation())).
                    SetUrlPosition(__GetPosition(node.GetLocation())).
                    SetUrl(url).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.CLASS, level, __GetName(node, true));
                foreach (var a_Context in node.Members.OfType<MethodBlockSyntax>())
                {
                    __Execute(a_Context, level + 1, context, url, false, isShowPrivate);
                }
                foreach (var a_Context in node.Members.OfType<PropertyBlockSyntax>())
                {
                    __Execute(a_Context, level + 1, context, url, isShowPrivate);
                }
                foreach (var a_Context in node.Members.OfType<FieldDeclarationSyntax>())
                {
                    __Execute(a_Context, level + 1, context, url, isShowPrivate);
                }
            }
        }

        private static void __Execute(EnumBlockSyntax node, int level, atom.Trace context, string url, bool isShowPrivate)
        {
            if (__IsEnabled(node.EnumStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetType(node.EnumStatement.Modifiers, "Enum")).
                    SetCommentHint(HINT.DATA_TYPE).
                    SetUrlLine(__GetLine(node.GetLocation())).
                    SetUrlPosition(__GetPosition(node.GetLocation())).
                    SetUrl(url).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.CLASS, level, __GetName(node, true));
                foreach (var a_Context in node.Members.OfType<EnumMemberDeclarationSyntax>())
                {
                    context.
                        SetComment("Integer").
                        SetCommentHint(HINT.DATA_TYPE).
                        SetUrlLine(__GetLine(a_Context.GetLocation())).
                        SetUrlPosition(__GetPosition(a_Context.GetLocation())).
                        SetUrl(url).
                        Send(NAME.SOURCE.PREVIEW, NAME.TYPE.INFO, level + 1, a_Context.Identifier.ValueText);
                }
            }
        }

        private static void __Execute(StructureBlockSyntax node, int level, atom.Trace context, string url, bool isShowPrivate)
        {
            if (__IsEnabled(node.StructureStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetType(node.StructureStatement.Modifiers, "Struct")).
                    SetCommentHint(HINT.DATA_TYPE).
                    SetUrlLine(__GetLine(node.GetLocation())).
                    SetUrlPosition(__GetPosition(node.GetLocation())).
                    SetUrl(url).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.CLASS, level, __GetName(node, true));
                foreach (var a_Context in node.Members.OfType<MethodBlockSyntax>())
                {
                    __Execute(a_Context, level + 1, context, url, false, isShowPrivate);
                }
                foreach (var a_Context in node.Members.OfType<PropertyBlockSyntax>())
                {
                    __Execute(a_Context, level + 1, context, url, isShowPrivate);
                }
                foreach (var a_Context in node.Members.OfType<FieldDeclarationSyntax>())
                {
                    __Execute(a_Context, level + 1, context, url, isShowPrivate);
                }
            }
        }

        private static void __Execute(MethodBlockSyntax node, int level, atom.Trace context, string url, bool isFullName, bool isShowPrivate)
        {
            if (__IsEnabled(node.BlockStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetComment(node)).
                    SetCommentHint(HINT.DATA_TYPE).
                    SetUrlLine(__GetLine(node.GetLocation())).
                    SetUrlPosition(__GetPosition(node.GetLocation())).
                    SetUrl(url).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.FUNCTION, level, __GetName(node, isFullName));
            }
        }

        private static void __Execute(PropertyBlockSyntax node, int level, atom.Trace context, string url, bool isShowPrivate)
        {
            if (__IsEnabled(node.PropertyStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetComment(node)).
                    SetCommentHint(HINT.DATA_TYPE).
                    SetUrlLine(__GetLine(node.GetLocation())).
                    SetUrlPosition(__GetPosition(node.GetLocation())).
                    SetUrl(url).
                    SetValue(node.PropertyStatement.Initializer?.Value?.ToString()).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.PARAMETER, level, node.PropertyStatement?.Identifier.ValueText);
            }
        }

        private static void __Execute(FieldDeclarationSyntax node, int level, atom.Trace context, string url, bool isShowPrivate)
        {
            if (__IsEnabled(node.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetComment(node)).
                    SetCommentHint(HINT.DATA_TYPE).
                    SetUrlLine(__GetLine(node.GetLocation())).
                    SetUrlPosition(__GetPosition(node.GetLocation())).
                    SetUrl(url).
                    SetValue(node.Declarators.First()?.Initializer?.Value?.ToString()).
                    Send(NAME.SOURCE.PREVIEW, NAME.TYPE.VARIABLE, level, node.Declarators.First()?.Names.First()?.ToString());
            }
        }

        private static bool __IsEnabled(SyntaxTokenList value, bool isShowPrivate)
        {
            if (GetState() == STATE.CANCEL)
            {
                return false;
            }
            if (isShowPrivate == false)
            {
                var a_Context = value.ToString();
                if (string.IsNullOrEmpty(a_Context) == false)
                {
                    if (a_Context.Contains("Private"))
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        private static string __GetType(SyntaxTokenList node, string typeName)
        {
            if (node != null)
            {
                return node.ToString().Trim() + " " + typeName;
            }
            return typeName;
        }

        internal static string __GetArraySize(IEnumerable value)
        {
            var a_Result = 0;
            foreach (var a_Context in value)
            {
                a_Result++;
            }
            return "[[Found]]: " + a_Result.ToString();
        }

        private static string __GetSeverity(Diagnostic node)
        {
            switch (node.Severity)
            {
                case DiagnosticSeverity.Hidden: return NAME.TYPE.DEBUG;
                case DiagnosticSeverity.Info: return NAME.TYPE.INFO;
                case DiagnosticSeverity.Warning: return NAME.TYPE.WARNING;
                case DiagnosticSeverity.Error: return NAME.TYPE.ERROR;
            }
            return NAME.TYPE.INFO;
        }

        private static string __GetName(SyntaxNode node, bool isFullName)
        {
            var a_Result = "";
            var a_Context = node;
            while (a_Context != null)
            {
                if (isFullName)
                {
                    if ((a_Context is NamespaceStatementSyntax) && (((a_Context as NamespaceStatementSyntax).Name as IdentifierNameSyntax) != null))
                    {
                        a_Result = ((a_Context as NamespaceStatementSyntax).Name as IdentifierNameSyntax).Identifier.ValueText + "." + a_Result;
                    }
                    if (a_Context is ClassBlockSyntax)
                    {
                        a_Result = (a_Context as ClassBlockSyntax).ClassStatement.Identifier.ValueText + (string.IsNullOrEmpty(a_Result) ? "" : ("." + a_Result));
                    }
                }
                else
                {
                    if ((a_Context is ClassBlockSyntax) && string.IsNullOrEmpty(a_Result))
                    {
                        a_Result = (a_Context as ClassBlockSyntax).ClassStatement.Identifier.ValueText;
                    }
                }
                if (a_Context is MethodBlockSyntax)
                {
                    a_Result = ((a_Context as MethodBlockSyntax).BlockStatement as MethodStatementSyntax)?.Identifier.ValueText + __GetParams(a_Context as MethodBlockSyntax);
                }
                if (a_Context is PropertyBlockSyntax)
                {
                    a_Result = (a_Context as PropertyBlockSyntax).PropertyStatement.Identifier.ValueText;
                }
                if (a_Context is EnumBlockSyntax)
                {
                    a_Result = (a_Context as EnumBlockSyntax).EnumStatement.Identifier.ValueText;
                }
                if (a_Context is StructureBlockSyntax)
                {
                    a_Result = (a_Context as StructureBlockSyntax).StructureStatement.Identifier.ValueText;
                }
                {
                    a_Context = a_Context.Parent;
                }
            }
            return a_Result;
        }

        private static string __GetParams(MethodBlockBaseSyntax node)
        {
            var a_Result = "";
            var a_Context = "";
            foreach (var a_Context1 in node.BlockStatement?.ParameterList.Parameters)
            {
                if (a_Context1.AsClause != null)
                {
                    a_Result += a_Context;
                    a_Result += a_Context1.AsClause?.Type?.ToString() + " ";
                    a_Result += a_Context1.Identifier?.ToString();
                    a_Context = ", ";
                }
            }
            return "(" + a_Result + ")";
        }

        private static string __GetComment(MethodBlockSyntax node)
        {
            var a_Context = node.BlockStatement as MethodStatementSyntax;
            if (a_Context?.AsClause != null)
            {
                return __GetType(a_Context.Modifiers, a_Context?.AsClause?.Type.ToString());
            }
            return __GetType(a_Context.Modifiers, "Void");
        }

        private static string __GetComment(PropertyBlockSyntax node)
        {
            var a_Context = node.PropertyStatement;
            if (a_Context?.AsClause != null)
            {
                return __GetType(a_Context.Modifiers, (a_Context.AsClause as SimpleAsClauseSyntax)?.Type.ToString());
            }
            return __GetType(a_Context.Modifiers, "Void");
        }

        private static string __GetComment(FieldDeclarationSyntax node)
        {
            var a_Context = node.Declarators.First();
            if (a_Context?.AsClause != null)
            {
                return __GetType(node.Modifiers, (a_Context.AsClause as SimpleAsClauseSyntax)?.Type.ToString());
            }
            return __GetType(node.Modifiers, "Void");
        }

        private static int __GetLine(Location node)
        {
            if (node.Kind != LocationKind.None)
            {
                return node.GetLineSpan().StartLinePosition.Line + 1;
            }
            return 0;
        }

        private static int __GetPosition(Location node)
        {
            if (node.Kind != LocationKind.None)
            {
                return node.GetLineSpan().StartLinePosition.Character + 1;
            }
            return 0;
        }
    };
}
