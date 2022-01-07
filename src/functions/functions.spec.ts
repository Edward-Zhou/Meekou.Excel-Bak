//import * as OfficeAddinMock from "office-addin-mock";
import { describe, it } from "mocha";
import * as assert from "assert";
import { add } from "./functions";
describe("test init", () => {
  it("test 1", () => {
    var result = add(1, 2);
    assert.strictEqual(603, result);
  });
});
