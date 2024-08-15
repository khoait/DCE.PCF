import Handlebars from "handlebars";

/**
 * Getting the variables from the Handlebars template.
 * Supports helpers too.
 * @param input
 */
export function getHandlebarsVariables(input: string): string[] {
  const ast = Handlebars.parseWithoutProcessing(input);

  return ast.body
    .filter(({ type }: hbs.AST.Statement) => type === "MustacheStatement")
    .map((statement: hbs.AST.Statement) => {
      const moustacheStatement: hbs.AST.MustacheStatement = statement as hbs.AST.MustacheStatement;
      const paramsExpressionList = moustacheStatement.params as hbs.AST.PathExpression[];
      const pathExpression = moustacheStatement.path as hbs.AST.PathExpression;

      return paramsExpressionList[0]?.original || pathExpression.original;
    });
}
