function writeRefinedContent_(content){
const doc = DocumentApp.getActiveDocument()
const body = doc.getBody()
body.appendParagraph("\n--- AI Refined Content ---\n")
body.appendParagraph(content)
return "Content added to document."
}