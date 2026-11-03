/* Small ZIP reader/writer for local PPTX editing. No file leaves the browser. */
window.PptZip = (() => {
  const decoder = new TextDecoder(), encoder = new TextEncoder();
  const crcTable = Array.from({length:256},(_,i)=>{let c=i;for(let j=0;j<8;j++)c=(c&1)?0xedb88320^(c>>>1):c>>>1;return c>>>0});
  function crc32(bytes){let c=0xffffffff;for(const b of bytes)c=crcTable[(c^b)&255]^(c>>>8);return (c^0xffffffff)>>>0}
  async function read(buffer){
    const bytes=new Uint8Array(buffer),view=new DataView(buffer);let end=-1;
    for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(view.getUint32(i,true)===0x06054b50){end=i;break}
    if(end<0)throw new Error('PPTX ZIP 형식을 읽을 수 없습니다.');
    const total=view.getUint16(end+10,true);let p=view.getUint32(end+16,true),entries=new Map(),unpacked=0;
    for(let i=0;i<total;i++){
      if(view.getUint32(p,true)!==0x02014b50)throw new Error('손상된 ZIP 목록입니다.');
      const method=view.getUint16(p+10,true),size=view.getUint32(p+20,true),rawSize=view.getUint32(p+24,true),nl=view.getUint16(p+28,true),el=view.getUint16(p+30,true),cl=view.getUint16(p+32,true),offset=view.getUint32(p+42,true);
      unpacked+=rawSize;if(unpacked>100*1024*1024)throw new Error('압축 해제 용량이 너무 큽니다. 100MB 이하 파일을 사용하세요.');
      const name=decoder.decode(bytes.subarray(p+46,p+46+nl)),dataStart=offset+30+view.getUint16(offset+26,true)+view.getUint16(offset+28,true),compressed=bytes.slice(dataStart,dataStart+size);let data;
      if(method===0)data=compressed;else if(method===8){
        if(!window.DecompressionStream)throw new Error('최신 Chrome 또는 Edge에서 이용해주세요.');
        const stream=new Blob([compressed]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
        data=new Uint8Array(await new Response(stream).arrayBuffer());
      }else throw new Error('지원하지 않는 압축 방식입니다.');
      if(data.length!==rawSize||crc32(data)!==view.getUint32(p+16,true))throw new Error('파일 무결성 확인에 실패했습니다.');
      entries.set(name,data);p+=46+nl+el+cl;
    }
    return entries;
  }
  function write(entries){
    let locals=[],centrals=[],offset=0;
    for(const [name,data] of entries){const n=encoder.encode(name),crc=crc32(data),local=new Uint8Array(30+n.length),lv=new DataView(local.buffer);lv.setUint32(0,0x04034b50,true);lv.setUint16(4,20,true);lv.setUint16(6,0x800,true);lv.setUint32(14,crc,true);lv.setUint32(18,data.length,true);lv.setUint32(22,data.length,true);lv.setUint16(26,n.length,true);local.set(n,30);locals.push(local,data);
      const central=new Uint8Array(46+n.length),cv=new DataView(central.buffer);cv.setUint32(0,0x02014b50,true);cv.setUint16(4,20,true);cv.setUint16(6,20,true);cv.setUint16(8,0x800,true);cv.setUint32(16,crc,true);cv.setUint32(20,data.length,true);cv.setUint32(24,data.length,true);cv.setUint16(28,n.length,true);cv.setUint32(42,offset,true);central.set(n,46);centrals.push(central);offset+=local.length+data.length;
    }
    const centralSize=centrals.reduce((s,b)=>s+b.length,0),end=new Uint8Array(22),ev=new DataView(end.buffer);ev.setUint32(0,0x06054b50,true);ev.setUint16(8,entries.size,true);ev.setUint16(10,entries.size,true);ev.setUint32(12,centralSize,true);ev.setUint32(16,offset,true);return new Blob([...locals,...centrals,end],{type:'application/vnd.openxmlformats-officedocument.presentationml.presentation'});
  }
  return {read,write,text:(entries,name)=>{if(!entries.has(name))throw new Error('PPTX 구성요소가 없습니다: '+name);return decoder.decode(entries.get(name))},setText:(entries,name,text)=>entries.set(name,encoder.encode(text))};
})();
