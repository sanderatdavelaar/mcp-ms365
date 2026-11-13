// Body voor een conceptantwoord: de tekst van het antwoord boven het geciteerde bericht dat Graph in het
// concept zette. Zonder citaat blijft het platte tekst.
const esc = (s: string) => s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

export function replyBody(text: string, quotedHtml?: string): { contentType: "HTML" | "Text"; content: string } {
  if (!quotedHtml || !quotedHtml.trim()) return { contentType: "Text", content: text };
  const reply = `<div>${esc(text).replace(/\r?\n/g, "<br>")}</div><br>`;
  const body = /<body[^>]*>/i.exec(quotedHtml);
  if (!body) return { contentType: "HTML", content: reply + quotedHtml };
  const at = body.index + body[0].length;
  return { contentType: "HTML", content: quotedHtml.slice(0, at) + reply + quotedHtml.slice(at) };
}
