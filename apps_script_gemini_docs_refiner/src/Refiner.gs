function refineDocument(userPrompt){
if(!userPrompt) return "Enter instructions."
const text = getDocumentText_()
const prompt = `
You are a writing assistant.

Document:
${text.substring(0,12000)}

Instruction:
${userPrompt}

Return only transformed content.
`
let result
try{
result = callGemini(prompt,"")
}catch(e){
return "Gemini error: "+e
}
return writeRefinedContent_(result)
}