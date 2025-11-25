import { GoogleGenAI } from "@google/genai";

const SYSTEM_INSTRUCTION = `
You are a world-class "Excel Auto-Edit Engineer". You are an expert in Python, specifically the \`pandas\` and \`openpyxl\` libraries.

Your goal is to generate Python code to manipulate an Excel file based strictly on the instructions found in cell "A1" of the first sheet.

**Context:**
- The input file is located at \`input.xlsx\`.
- The output file must be saved to \`output.xlsx\`.
- The Python environment is running inside a browser (Pyodide).
- You have access to \`pandas\` and \`openpyxl\`.
- **Language**: The user is Japanese. The instruction in A1 will likely be in Japanese.
- **Logging**: Use \`print()\` statements frequently to report progress. **These messages MUST be in Japanese.**

**Process:**
1.  Receive the content of cell A1 (the instruction) and a summary of the dataframe columns.
2.  Generate a complete, robust Python script to perform the task.
3.  **A1 Handling**: Unless explicitly told to delete it, keep the instruction in A1 or overwrite it only if the instruction implies replacing the whole sheet. If unsure, leave it.
4.  **Error Handling**: Wrap critical logic in try-except blocks.
5.  **Output**: Return ONLY the Python code block. Do not include conversational text outside the code block.

**Code Requirements:**
- Import necessary libraries (\`pandas\`, \`openpyxl\`).
- Read \`input.xlsx\`.
- Execute the logic described in the A1 instruction.
- Save the result to \`output.xlsx\`.
- Use \`print()\` statements to log progress (in Japanese).
`;

export const generateExcelEditCode = async (
  a1Instruction: string,
  columns: string[]
): Promise<string> => {
  const apiKey = process.env.API_KEY;
  if (!apiKey) {
    throw new Error("API Key is missing.");
  }

  const ai = new GoogleGenAI({ apiKey });

  const prompt = `
    The user has uploaded an Excel file.
    
    **Cell A1 Instruction:** "${a1Instruction}"
    
    **Sheet Columns:** ${columns.join(', ')}
    
    Generate the Python script to execute this instruction.
    Input file: 'input.xlsx'
    Output file: 'output.xlsx'
    
    Return only valid Python code wrapped in \`\`\`python ... \`\`\`.
  `;

  try {
    const response = await ai.models.generateContent({
      model: 'gemini-2.5-flash',
      contents: prompt,
      config: {
        systemInstruction: SYSTEM_INSTRUCTION,
        temperature: 0.2, // Low temperature for precise code generation
      }
    });

    const text = response.text;
    if (!text) throw new Error("No response from Gemini");

    // Extract code block
    const codeMatch = text.match(/```python([\s\S]*?)```/);
    if (codeMatch && codeMatch[1]) {
      return codeMatch[1].trim();
    }
    // Fallback if no block found but text exists (rare)
    return text.replace(/```python/g, '').replace(/```/g, '').trim();

  } catch (error: any) {
    console.error("Gemini API Error:", error);
    throw new Error(`Failed to generate code: ${error.message}`);
  }
};