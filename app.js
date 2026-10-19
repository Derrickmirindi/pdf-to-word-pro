const $=id=>document.getElementById(id);
let url;
try {
 const [pdfjs,d]=await Promise.all([import('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.min.mjs'),import('https://cdn.jsdelivr.net/npm/docx@9.1.1/+esm')]);
 pdfjs.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.10.38/pdf.worker.min.mjs';
 $('convert').disabled=false;$('convert').textContent='Convert to Word';
 $('convert').onclick=async()=>{
  const file=$('file').files[0];if(!file){$('status').textContent='Please select a PDF first.';return;}
  if(file.size>50*1024*1024){$('status').textContent='Please use a PDF smaller than 50 MB.';return;}
  $('convert').disabled=true;$('download').hidden=true;let pdf;
  try{
   pdf=await pdfjs.getDocument({data:new Uint8Array(await file.arrayBuffer()),isEvalSupported:false}).promise;
   if(pdf.numPages>150)throw Error('Please split this PDF into documents of 150 pages or fewer.');
   const sections=[];let empty=0;
   for(let n=1;n<=pdf.numPages;n++){
    $('status').textContent=`Converting page ${n} of ${pdf.numPages}…`;
    await new Promise(r=>setTimeout(r,0));
    const page=await pdf.getPage(n),view=page.getViewport({scale:1}),children=[];
    if($('mode').value==='visual'){
     const scale=Math.min(2,2400/Math.max(view.width,view.height)),v=page.getViewport({scale});
     const canvas=document.createElement('canvas');canvas.width=Math.ceil(v.width);canvas.height=Math.ceil(v.height);
     await page.render({canvasContext:canvas.getContext('2d'),viewport:v}).promise;
     const blob=await new Promise(r=>canvas.toBlob(r,'image/jpeg',0.92));if(!blob)throw Error('Could not render page.');
     children.push(new d.Paragraph({spacing:{before:0,after:0},children:[new d.ImageRun({type:'jpg',data:new Uint8Array(await blob.arrayBuffer()),transformation:{width:view.width*96/72*.97,height:view.height*96/72*.97}})]}));canvas.width=canvas.height=0;
    }else{
     const content=await page.getTextContent();let line='',lastY=null;
     const flush=()=>{if(line.trim())children.push(new d.Paragraph({children:[new d.TextRun({text:line.trim(),size:22})]}));line='';};
     for(const item of content.items){if(!('str' in item))continue;const y=item.transform[5];if(lastY!==null&&Math.abs(y-lastY)>4)flush();line+=item.str+' ';lastY=y;if(item.hasEOL)flush();}flush();
     if(!children.length){empty++;children.push(new d.Paragraph('[No extractable text on this page. Try visual mode for scans.]'));}
    }
    sections.push({properties:{type:d.SectionType.NEXT_PAGE,page:{size:{width:Math.round(view.width*20),height:Math.round(view.height*20)},margin:{top:0,bottom:0,left:$('mode').value==='visual'?0:720,right:$('mode').value==='visual'?0:720}}},children});page.cleanup();
   }
   const blob=await d.Packer.toBlob(new d.Document({sections}));if(url)URL.revokeObjectURL(url);url=URL.createObjectURL(blob);
   $('download').href=url;$('download').download=file.name.replace(/\.pdf$/i,'')+'.docx';$('download').hidden=false;
   $('status').textContent=`Done. ${pdf.numPages} pages converted.${empty?' Some pages had no extractable text.':''} Click Download Word document.`;
  }catch(e){$('status').textContent='Conversion failed: '+e.message;}finally{if(pdf)await pdf.destroy();$('convert').disabled=false;}
 };
}catch(e){$('status').textContent='Could not load converter dependencies. Check your internet connection and refresh. '+e.message;}
