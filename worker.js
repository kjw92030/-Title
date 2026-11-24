const BASE = 'https://cdn.jsdelivr.net/pyodide/v0.27.7/full/';
let py;
let chain = Promise.resolve();
async function loadWorkbook(payload) {
  postMessage({type:'status',text:payload.label+' Excel 읽기 중…'});
  py.globals.set('dataset_input',payload.dataset);
  py.globals.set('raw_input',new Uint8Array(payload.buffer));
  await py.runPythonAsync('raw_bytes = raw_input.to_py()');
  const result=await py.runPythonAsync("load_dataset(dataset_input, raw_bytes)");
  await py.runPythonAsync('del raw_input, raw_bytes');
  return result;
}
async function handle({id,type,payload}) {
  try {
    let result;
    if(type==='init') {
      postMessage({type:'status',text:'Python 분석 환경 준비 중…'});
      importScripts(BASE+'pyodide.js');
      py=await loadPyodide({indexURL:BASE});
      await py.loadPackage(['numpy','micropip']);
      await py.runPythonAsync("import micropip\nawait micropip.install('openpyxl==3.1.5')");
      const response=await fetch('./model.py');
      if(!response.ok) throw Error('model.py를 불러올 수 없습니다.');
      await py.runPythonAsync(await response.text());
      result=await loadWorkbook(payload);
    } else if(type==='load') result=await loadWorkbook(payload);
    else if(type==='explore') {
      py.globals.set('dataset_input',payload.dataset);
      py.globals.set('sheet_input',payload.sheet);
      py.globals.set('row_input',payload.row||'');
      result=py.runPython('explore(dataset_input, sheet_input, row_input)');
    } else {
      py.globals.set('scenario_input',JSON.stringify(payload));
      result=py.runPython('simulate(scenario_input)');
    }
    postMessage({id,ok:true,result:JSON.parse(result)});
  } catch(error) { postMessage({id,ok:false,error:String(error)}); }
}
onmessage=e=>{ chain=chain.then(()=>handle(e.data)); };
