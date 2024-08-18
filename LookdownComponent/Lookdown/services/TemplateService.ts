import Handlebars from "handlebars";
import { getCurrentRecord } from "./DataverseService";

export function getFetchTemplateString(fetchXml: string, customFilter?: string | null, additionalColumns?: string[]) {
  if (fetchXml === "") return fetchXml;

  if (!customFilter && additionalColumns?.length === 0) return fetchXml;

  let templateString = fetchXml;

  const parser = new DOMParser();
  const doc = parser.parseFromString(fetchXml, "application/xml");
  const entity = doc.querySelector("entity");

  if (customFilter) {
    // create customFilterElement from customFilter
    const customFilterEl = parser.parseFromString(customFilter, "text/xml");
    entity?.appendChild(customFilterEl.documentElement);
  }

  if (additionalColumns?.length) {
    // get unique list of additionalColumns
    additionalColumns = [...new Set(additionalColumns)];

    // get all attribute names from fetchXml
    const attributes = doc.querySelectorAll("attribute");
    const attributeNames = Array.from(attributes).map((a) => a.getAttribute("name"));

    // get list of strings from additionalColumns not in attributeNames
    const addColumns = additionalColumns.filter((c) => !attributeNames.includes(c));

    addColumns.forEach((c) => {
      const attributeEl = parser.parseFromString(`<attribute name='${c}' />`, "text/xml");
      entity?.appendChild(attributeEl.documentElement);
    });
  }

  templateString = new XMLSerializer().serializeToString(doc);

  const fetchXmlTemplate = Handlebars.compile(templateString);
  return fetchXmlTemplate(getCurrentRecord());
}

export function getCustomFilterString(customFilter: string) {
  if (customFilter === "") return customFilter;

  const parser = new DOMParser();
  const doc = parser.parseFromString(customFilter, "application/xml");
  const filter = doc.querySelector("filter");

  if (!filter) return customFilter;

  const filterTemplate = Handlebars.compile(customFilter);
  return filterTemplate(getCurrentRecord());
}

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
