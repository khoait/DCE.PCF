import Handlebars from "handlebars";

export const getHBVars = (value: string) => {
  const ast = Handlebars.parse(value);
  const keys = [];

  for (const i in ast.body) {
    if (ast.body[i].type === "MustacheStatement") {
      const statement = ast.body[i] as hbs.AST.MustacheStatement;
      if (statement.path.type === "PathExpression") {
        const varName = (statement.path as hbs.AST.PathExpression).original;
        const attr = varName.split(".").at(0);
        if (attr) {
          keys.push(attr);
        }
      }
    }
  }
  return Array.from(new Set<string>(keys));
};
