/**
 * Encodeert een ID uit tool-argumenten als één padsegment voor Microsoft Graph.
 * Zonder encoding kan een "ID" als `../messages/<id>` een andere Graph-resource
 * raken dan de tool bedoelt (bv. delete_event dat een mail verwijdert).
 * Lege ID's en `.`/`..` worden geweigerd.
 */
export function gid(id: string | undefined): string {
  if (typeof id !== "string" || id.length === 0 || id === "." || id === "..") {
    throw new Error("Ongeldig ID");
  }
  return encodeURIComponent(id);
}
