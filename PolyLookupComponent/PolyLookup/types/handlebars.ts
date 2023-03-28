declare module "handlebars/lib/handlebars" {
  // Re-export the types from the existing handlebars module
  import Handlebars = require("handlebars");
  export default Handlebars;
}
