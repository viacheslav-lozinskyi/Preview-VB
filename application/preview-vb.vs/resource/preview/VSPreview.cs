
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.VisualBasic;
using Microsoft.CodeAnalysis.VisualBasic.Syntax;
using System.Collections;
using System.IO;
using System.Linq;

namespace resource.preview
{
    internal class VSPreview : extension.AnyPreview
    {
        internal class HINT
        {
            public static string DATA_TYPE = "[[[Data Type]]]";
            public static string METHOD_TYPE = "[[[Method Type]]]";
        }

        protected override void _Execute(atom.Trace context, int level, string url, string file)
        {
            var a_Context = VisualBasicSyntaxTree.ParseText(File.ReadAllText(file)).WithFilePath(file).GetRoot();
            var a_IsFound = GetProperty(NAME.PROPERTY.DEBUGGING_SHOW_PRIVATE, true) != 0;
            if (a_Context == null)
            {
                return;
            }
            {
                context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.HEADER, level, "[[[Info]]]");
                {
                    context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level + 1, "[[[File Name]]]", url);
                    context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level + 1, "[[[File Size]]]", a_Context.GetText().Length.ToString());
                    context.Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level + 1, "[[[Language]]]", a_Context.Language);
                }
            }
            if (a_Context.DescendantNodes().OfType<ImportsStatementSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<ImportsStatementSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FOLDER, level, "[[[Dependencies]]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<ImportsStatementSyntax>())
                {
                    __Execute(context, level + 1, a_Context1, file);
                }
            }
            if (a_Context.DescendantNodes().OfType<ClassBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<ClassBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FOLDER, level, "[[[Classes]]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<ClassBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context1, file, a_IsFound);
                }
            }
            if (a_Context.DescendantNodes().OfType<StructureBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<StructureBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FOLDER, level, "[[[Structs]]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<StructureBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context1, file, a_IsFound);
                }
            }
            if (a_Context.DescendantNodes().OfType<EnumBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<EnumBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FOLDER, level, "[[[Enums]]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<EnumBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context1, file, a_IsFound);
                }
            }
            if (a_Context.DescendantNodes().OfType<MethodBlockSyntax>().Any())
            {
                context.
                    SetComment(__GetArraySize(a_Context.DescendantNodes().OfType<MethodBlockSyntax>())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FOLDER, level, "[[[Functions]]]");
                foreach (var a_Context1 in a_Context.DescendantNodes().OfType<MethodBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context1, file, true, a_IsFound);
                }
            }
            if (a_Context.GetDiagnostics().Any())
            {
                context.
                    SendPreview(NAME.EVENT.ERROR, url).
                    SetComment(__GetArraySize(a_Context.GetDiagnostics())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.ERROR, level, "[[[Diagnostics]]]");
                foreach (var a_Context1 in a_Context.GetDiagnostics())
                {
                    __Execute(context, level + 1, a_Context1, file);
                }
            }
        }

        private static void __Execute(atom.Trace context, int level, Diagnostic data, string file)
        {
            context.
                SetUrl(file, __GetLine(data.Location), __GetPosition(data.Location)).
                SetUrlInfo("https://www.bing.com/search?q=" + data.Id).
                Send(NAME.SOURCE.PREVIEW, __GetType(data), level, data.Descriptor?.MessageFormat?.ToString());
        }

        private static void __Execute(atom.Trace context, int level, ImportsStatementSyntax data, string file)
        {
            context.
                SetComment("Imports", HINT.DATA_TYPE).
                SetUrl(file, __GetLine(data.GetLocation()), __GetPosition(data.GetLocation())).
                Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FILE, level, data.ImportsClauses.ToString());
        }

        private static void __Execute(atom.Trace context, int level, ClassBlockSyntax data, string file, bool isShowPrivate)
        {
            if (__IsEnabled(data.BlockStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment("Class", HINT.DATA_TYPE).
                    SetUrl(file, __GetLine(data.GetLocation()), __GetPosition(data.GetLocation())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.CLASS, level, __GetName(data, true));
                foreach (var a_Context in data.Members.OfType<MethodBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context, file, false, isShowPrivate);
                }
                foreach (var a_Context in data.Members.OfType<PropertyBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context, file, isShowPrivate);
                }
                foreach (var a_Context in data.Members.OfType<FieldDeclarationSyntax>())
                {
                    __Execute(context, level + 1, a_Context, file, isShowPrivate);
                }
            }
        }

        private static void __Execute(atom.Trace context, int level, EnumBlockSyntax data, string file, bool isShowPrivate)
        {
            if (__IsEnabled(data.EnumStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetType(data.EnumStatement.Modifiers, "Enum"), HINT.DATA_TYPE).
                    SetUrl(file, __GetLine(data.GetLocation()), __GetPosition(data.GetLocation())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.CLASS, level, __GetName(data, true));
                foreach (var a_Context in data.Members.OfType<EnumMemberDeclarationSyntax>())
                {
                    context.
                        SetComment("Integer", HINT.DATA_TYPE).
                        SetUrl(file, __GetLine(a_Context.GetLocation()), __GetPosition(a_Context.GetLocation())).
                        Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level + 1, a_Context.Identifier.ValueText);
                }
            }
        }

        private static void __Execute(atom.Trace context, int level, StructureBlockSyntax data, string file, bool isShowPrivate)
        {
            if (__IsEnabled(data.StructureStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetType(data.StructureStatement.Modifiers, "Struct"), HINT.DATA_TYPE).
                    SetUrl(file, __GetLine(data.GetLocation()), __GetPosition(data.GetLocation())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.CLASS, level, __GetName(data, true));
                foreach (var a_Context in data.Members.OfType<MethodBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context, file, false, isShowPrivate);
                }
                foreach (var a_Context in data.Members.OfType<PropertyBlockSyntax>())
                {
                    __Execute(context, level + 1, a_Context, file, isShowPrivate);
                }
                foreach (var a_Context in data.Members.OfType<FieldDeclarationSyntax>())
                {
                    __Execute(context, level + 1, a_Context, file, isShowPrivate);
                }
            }
        }

        private static void __Execute(atom.Trace context, int level, MethodBlockSyntax data, string file, bool isFullName, bool isShowPrivate)
        {
            if (__IsEnabled(data.BlockStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetComment(data), HINT.DATA_TYPE).
                    SetUrl(file, __GetLine(data.GetLocation()), __GetPosition(data.GetLocation())).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.FUNCTION, level, __GetName(data, isFullName));
            }
        }

        private static void __Execute(atom.Trace context, int level, PropertyBlockSyntax data, string file, bool isShowPrivate)
        {
            if (__IsEnabled(data.PropertyStatement.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetComment(data), HINT.DATA_TYPE).
                    SetUrl(file, __GetLine(data.GetLocation()), __GetPosition(data.GetLocation())).
                    SetValue(data.PropertyStatement.Initializer?.Value?.ToString()).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.PARAMETER, level, data.PropertyStatement?.Identifier.ValueText);
            }
        }

        private static void __Execute(atom.Trace context, int level, FieldDeclarationSyntax data, string file, bool isShowPrivate)
        {
            if (__IsEnabled(data.Modifiers, isShowPrivate))
            {
                context.
                    SetComment(__GetComment(data), HINT.DATA_TYPE).
                    SetUrl(file, __GetLine(data.GetLocation()), __GetPosition(data.GetLocation())).
                    SetValue(data.Declarators.First()?.Initializer?.Value?.ToString()).
                    Send(NAME.SOURCE.PREVIEW, NAME.EVENT.VARIABLE, level, data.Declarators.First()?.Names.First()?.ToString());
            }
        }

        private static bool __IsEnabled(SyntaxTokenList data, bool isShowPrivate)
        {
            if (GetState() == NAME.STATE.CANCEL)
            {
                return false;
            }
            if (isShowPrivate == false)
            {
                var a_Context = data.ToString();
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

        private static string __GetType(SyntaxTokenList data, string typeName)
        {
            if (data != null)
            {
                return data.ToString().Trim() + " " + typeName;
            }
            return typeName;
        }

        private static string __GetType(Diagnostic data)
        {
            switch (data.Severity)
            {
                case DiagnosticSeverity.Hidden: return NAME.EVENT.DEBUG;
                case DiagnosticSeverity.Info: return NAME.EVENT.PARAMETER;
                case DiagnosticSeverity.Warning: return NAME.EVENT.WARNING;
                case DiagnosticSeverity.Error: return NAME.EVENT.ERROR;
            }
            return NAME.EVENT.PARAMETER;
        }

        internal static string __GetArraySize(IEnumerable data)
        {
            var a_Result = 0;
            foreach (var a_Context in data)
            {
                a_Result++;
            }
            return "[[[Found]]]: " + a_Result.ToString();
        }

        private static string __GetName(SyntaxNode data, bool isFullName)
        {
            var a_Result = "";
            var a_Context = data;
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

        private static string __GetParams(MethodBlockBaseSyntax data)
        {
            var a_Result = "";
            var a_Context = "";
            foreach (var a_Context1 in data.BlockStatement?.ParameterList.Parameters)
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

        private static string __GetComment(MethodBlockSyntax data)
        {
            var a_Context = data.BlockStatement as MethodStatementSyntax;
            if (a_Context?.AsClause != null)
            {
                return __GetType(a_Context.Modifiers, a_Context?.AsClause?.Type.ToString());
            }
            return __GetType(a_Context.Modifiers, "Void");
        }

        private static string __GetComment(PropertyBlockSyntax data)
        {
            var a_Context = data.PropertyStatement;
            if (a_Context?.AsClause != null)
            {
                return __GetType(a_Context.Modifiers, (a_Context.AsClause as SimpleAsClauseSyntax)?.Type.ToString());
            }
            return __GetType(a_Context.Modifiers, "Void");
        }

        private static string __GetComment(FieldDeclarationSyntax data)
        {
            var a_Context = data.Declarators.First();
            if (a_Context?.AsClause != null)
            {
                return __GetType(data.Modifiers, (a_Context.AsClause as SimpleAsClauseSyntax)?.Type.ToString());
            }
            return __GetType(data.Modifiers, "Void");
        }

        private static int __GetLine(Location data)
        {
            if (data.Kind != LocationKind.None)
            {
                return data.GetLineSpan().StartLinePosition.Line + 1;
            }
            return 0;
        }

        private static int __GetPosition(Location data)
        {
            if (data.Kind != LocationKind.None)
            {
                return data.GetLineSpan().StartLinePosition.Character + 1;
            }
            return 0;
        }
    };
}
