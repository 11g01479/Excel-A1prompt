// Define types for global Pyodide
declare global {
  interface Window {
    loadPyodide: (config: any) => Promise<any>;
    pyodideInstance: any;
  }
}

let pyodideReadyPromise: Promise<any> | null = null;

export const initPyodide = async (logCallback: (msg: string) => void) => {
  if (window.pyodideInstance) return window.pyodideInstance;

  if (!pyodideReadyPromise) {
    pyodideReadyPromise = (async () => {
      logCallback("Python実行環境(Pyodide)をロード中...");
      const pyodide = await window.loadPyodide({
        indexURL: "https://cdn.jsdelivr.net/pyodide/v0.25.1/full/",
      });
      
      logCallback("必要なライブラリ(pandas, openpyxl)をインストール中... これには時間がかかる場合があります。");
      await pyodide.loadPackage("micropip");
      const micropip = pyodide.pyimport("micropip");
      await micropip.install(["pandas", "openpyxl"]);
      
      logCallback("Python環境の準備が完了しました。");
      window.pyodideInstance = pyodide;
      return pyodide;
    })();
  }
  return pyodideReadyPromise;
};

export const extractA1AndColumns = async (
  file: File
): Promise<{ a1: string; columns: string[] }> => {
  const pyodide = window.pyodideInstance;
  if (!pyodide) throw new Error("Pyodide not initialized");

  const arrayBuffer = await file.arrayBuffer();
  pyodide.FS.writeFile("input_temp.xlsx", new Uint8Array(arrayBuffer));

  const script = `
import pandas as pd
import openpyxl

# Use openpyxl to get raw A1 value exactly
wb = openpyxl.load_workbook("input_temp.xlsx")
ws = wb.active
a1_val = ws['A1'].value

# Use pandas to get column names easily
df = pd.read_excel("input_temp.xlsx")
cols = list(df.columns)

# Return dictionary
import json
json.dumps({
    "a1": str(a1_val) if a1_val else "",
    "columns": [str(c) for c in cols]
})
`;

  const resultJson = await pyodide.runPythonAsync(script);
  return JSON.parse(resultJson);
};

export const runPythonTransformation = async (
  script: string,
  inputFile: File,
  logCallback: (msg: string) => void
): Promise<Blob> => {
  const pyodide = window.pyodideInstance;
  if (!pyodide) throw new Error("Pyodide not initialized");

  // Mount file
  logCallback("ファイルを仮想環境にマウント中...");
  const arrayBuffer = await inputFile.arrayBuffer();
  pyodide.FS.writeFile("input.xlsx", new Uint8Array(arrayBuffer));

  // Redirect stdout to log
  pyodide.setStdout({
    batched: (msg: string) => logCallback(`[Python] ${msg}`),
  });

  logCallback("生成されたPythonスクリプトを実行中...");
  try {
    await pyodide.runPythonAsync(script);
  } catch (err: any) {
    throw new Error(`Pythonランタイムエラー: ${err.message}`);
  }

  // Retrieve output
  logCallback("出力ファイルを確認中...");
  if (pyodide.FS.analyzePath("output.xlsx").exists) {
    const fileData = pyodide.FS.readFile("output.xlsx");
    return new Blob([fileData], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  } else {
    throw new Error("スクリプトは完了しましたが、'output.xlsx' が生成されませんでした。");
  }
};