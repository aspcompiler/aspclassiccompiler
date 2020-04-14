using System;
using System.Collections.Generic;
using System.Text;
using VB = Dlrsoft.VBScript.Parser;
using Microsoft.Scripting.Runtime;

using System.Dynamic;
#if USE35
using Microsoft.Scripting.Ast;
#else
using System.Linq.Expressions;
#endif

using Microsoft.Scripting.Utils;
using System.Reflection;
using System.IO;
using Dlrsoft.VBScript.Binders;
using Dlrsoft.VBScript.Compiler;
using Dlrsoft.VBScript.Runtime;
using Debug = System.Diagnostics.Debug;

namespace Dlrsoft.VBScript.Compiler
{
    /// <summary>
    /// The analysis phase. We don't do much except for generating the global functions, constants and variables table
    /// </summary>
    internal class VBScriptAnalyzer
    {
        public static void AnalyzeFile(Dlrsoft.VBScript.Parser.ScriptBlock block, AnalysisScope scope)
        {
            Debug.Assert(scope.IsModule);

            if (block.Statements != null)
            {
                foreach (VB.Statement s in block.Statements)
                {
                    AnalyzeStatement(s, scope);
                }
            }
        }

        private static void AnalyzeStatement(VB.Statement s, AnalysisScope scope)
        {
            if (s is VB.MethodDeclaration)
            {
                string methodName = ((VB.MethodDeclaration)s).Name.Name.ToLower();
                scope.FunctionTable.Add(methodName);
                ParameterExpression method = Expression.Parameter(typeof(object));
                scope.Names[methodName] = method;
            }
            else if (s is VB.LocalDeclarationStatement)
            {
                AnalyzeDeclarationExpr((VB.LocalDeclarationStatement)s, scope);
            }
            else if (s is VB.IfBlockStatement)
            {
                AnalyzeIfBlockStatement((VB.IfBlockStatement)s, scope);
            }
            else if (s is VB.SelectBlockStatement)
            {
                AnalyzeSelectBlockStatement((VB.SelectBlockStatement)s, scope);
            }
            else if (s is VB.BlockStatement)
            {
                AnalyzeBlockStatement((VB.BlockStatement)s, scope);
            }
        }

        private static void AnalyzeDeclarationExpr(VB.LocalDeclarationStatement stmt, AnalysisScope scope)
        {
            bool isConst = false;

            if (stmt.Modifiers != null)
            {
                if (stmt.Modifiers.ModifierTypes == VB.ModifierTypes.Const)
                {
                    isConst = true;
                }
            }

            foreach (VB.VariableDeclarator d in stmt.VariableDeclarators)
            {
                foreach (VB.VariableName v in d.VariableNames)
                {
                    string name = v.Name.Name.ToLower();

                    ParameterExpression p = Expression.Parameter(typeof(object), name);
                    scope.Names.Add(name, p);
                }
            }
        }

        private static void AnalyzeIfBlockStatement(VB.IfBlockStatement block, AnalysisScope scope)
        {
            AnalyzeBlockStatement(block, scope);
            if (block.ElseIfBlockStatements != null && block.ElseIfBlockStatements.Count > 0)
            {
                foreach (VB.ElseIfBlockStatement elseif in block.ElseIfBlockStatements)
                {
                    AnalyzeBlockStatement(elseif, scope);
                }
            }
            if (block.ElseBlockStatement != null)
            {
                AnalyzeBlockStatement(block.ElseBlockStatement, scope);
            }
        }

        private static void AnalyzeSelectBlockStatement(VB.SelectBlockStatement block, AnalysisScope scope)
        {
            AnalyzeBlockStatement(block, scope);
            if (block.CaseBlockStatements != null && block.CaseBlockStatements.Count > 0)
            {
                foreach (VB.CaseBlockStatement caseStmt in block.CaseBlockStatements)
                {
                    AnalyzeBlockStatement(caseStmt, scope);
                }
            }
            if (block.CaseElseBlockStatement != null)
            {
                AnalyzeBlockStatement(block.CaseElseBlockStatement, scope);
            }
        }

        public static void AnalyzeBlockStatement(VB.BlockStatement block, AnalysisScope scope)
        {
            if (block.Statements != null)
            {
                foreach (VB.Statement stmt in block.Statements)
                {
                    AnalyzeStatement(stmt, scope);
                }
            }
        }
    }
}
