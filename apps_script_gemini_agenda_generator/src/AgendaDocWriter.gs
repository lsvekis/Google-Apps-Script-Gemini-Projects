function createAgendaDoc_(agenda){

const doc = DocumentApp.create(
agenda.title || "AI Agenda"
)

const body = doc.getBody()

body.appendParagraph(
agenda.title || "Meeting Agenda"
)

body.appendParagraph(
"Objective: " + (agenda.objective || "")
)

;(agenda.sections || []).forEach(s=>{

body.appendParagraph(
s.topic + " (" + s.duration + ")"
)

body.appendParagraph(
s.details || ""
)

})

return doc.getUrl()

}