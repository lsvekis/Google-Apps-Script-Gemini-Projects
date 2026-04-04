function getDocumentText_(){
const doc = DocumentApp.getActiveDocument()
return doc.getBody().getText()
}