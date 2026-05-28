function generateAgenda(userPrompt){

const prompt = "Create structured meeting agenda JSON for: " + userPrompt

let result = callGemini(prompt,"")

result = result.replace(/```json/g,"")
.replace(/```/g,"")

const json = JSON.parse(result)

return createAgendaDoc_(json)

}