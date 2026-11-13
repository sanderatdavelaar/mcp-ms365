import { test } from "node:test";
import assert from "node:assert/strict";
import { replyBody } from "../dist/tools/reply-body.js";

test("het antwoord komt boven het geciteerde bericht te staan, als tekst", () => {
  const quoted = '<html><head><meta charset="utf-8"></head><body lang="NL"><hr><div><b>Van:</b> Ingrid</div><p>Is Paddock al gereed?</p></body></html>';
  const out = replyBody("Hoi Ingrid,\n\nNog niet <klaar> & getest.\n\nGroet,\nSander", quoted);
  assert.equal(out.contentType, "HTML");
  const i = out.content.indexOf("Hoi Ingrid,");
  assert.ok(i > out.content.indexOf('<body lang="NL">'), "het antwoord hoort binnen de body");
  assert.ok(i < out.content.indexOf("Is Paddock al gereed?"), "het antwoord hoort boven het citaat");
  assert.ok(out.content.includes("Nog niet &lt;klaar&gt; &amp; getest."), "tekens uit het antwoord zijn geen opmaak");
  assert.ok(out.content.includes("Hoi Ingrid,<br><br>Nog niet"), "regeleinden blijven regeleinden");
  assert.ok(out.content.includes("<b>Van:</b> Ingrid"), "het citaat blijft ongewijzigd");
});

test("zonder body-tag komt het antwoord vooraan", () => {
  const out = replyBody("Kort", "<div>origineel</div>");
  assert.equal(out.contentType, "HTML");
  assert.ok(out.content.indexOf("Kort") < out.content.indexOf("origineel"));
});

test("een origineel in platte tekst blijft onder het antwoord staan", () => {
  const out = replyBody("Hoi,\n\nPrima.", "\r\n________________________________\r\nVan: Ingrid\r\nIs het klaar?", "Text");
  assert.equal(out.contentType, "Text");
  assert.ok(out.content.startsWith("Hoi,\n\nPrima.\n\n"));
  assert.ok(out.content.includes("Van: Ingrid\r\nIs het klaar?"));
});

test("zonder citaat blijft het platte tekst, zoals voorheen", () => {
  assert.deepEqual(replyBody("Alleen dit", ""), { contentType: "Text", content: "Alleen dit" });
  assert.deepEqual(replyBody("Alleen dit", undefined), { contentType: "Text", content: "Alleen dit" });
});
