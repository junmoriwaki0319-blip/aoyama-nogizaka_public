#!/usr/bin/env node
const fs = require('fs');
const path = require('path');
const YahooFinance = require('yahoo-finance2').default;
const yf = new YahooFinance({ suppressNotices: ['yahooSurvey'] });

const COMPANIES = [
  // 放送
  {name:'フジ・メディアHD',ticker:'4676',category:'broadcasting'},
  {name:'日本テレビHD',ticker:'9404',category:'broadcasting'},
  {name:'TBS HD',ticker:'9401',category:'broadcasting'},
  {name:'テレビ朝日HD',ticker:'9409',category:'broadcasting'},
  {name:'テレビ東京HD',ticker:'9413',category:'broadcasting'},
  {name:'WOWOW',ticker:'4839',category:'broadcasting'},
  {name:'スカパーJSAT HD',ticker:'9412',category:'broadcasting'},
  {name:'日本BS放送',ticker:'9414',category:'broadcasting'},
  // 出版
  {name:'KADOKAWA',ticker:'9468',category:'publishing'},
  {name:'学研HD',ticker:'9470',category:'publishing'},
  {name:'インプレスHD',ticker:'9479',category:'publishing'},
  {name:'ゼンリン',ticker:'9474',category:'publishing'},
  {name:'メディアドゥ',ticker:'3678',category:'publishing'},
  {name:'イーブックイニシアティブジャパン',ticker:'3658',category:'publishing'},
  // プラットフォーム
  {name:'LINEヤフー',ticker:'4689',category:'platform'},
  {name:'note',ticker:'5243',category:'platform'},
  {name:'ユーザベース',ticker:'4966',category:'platform'},
  {name:'Gunosy',ticker:'6047',category:'platform'},
  {name:'はてな',ticker:'3930',category:'platform'},
  {name:'メドピア',ticker:'6095',category:'platform'},
  {name:'kubell',ticker:'4448',category:'platform'},
  {name:'じげん',ticker:'3679',category:'platform'},
  {name:'サイバーエージェント',ticker:'4751',category:'platform'},
  // 動画・映像
  {name:'U-NEXT HOLDINGS',ticker:'9418',category:'video'},
  {name:'IMAGICA GROUP',ticker:'6879',category:'video'},
  {name:'AOI TYO Holdings',ticker:'3975',category:'video'},
  {name:'東北新社',ticker:'2329',category:'video'},
  {name:'東宝',ticker:'9602',category:'video'},
  {name:'松竹',ticker:'9601',category:'video'},
  {name:'東映',ticker:'9605',category:'video'},
  {name:'東映アニメーション',ticker:'4816',category:'video'},
  {name:'IGポート',ticker:'3791',category:'video'},
  {name:'Jストリーム',ticker:'4308',category:'video'},
  {name:'ハピネット',ticker:'7552',category:'video'},
  {name:'クリーク・アンド・リバー社',ticker:'4763',category:'video'},
  {name:'デジタルハーツHD',ticker:'3676',category:'video'},
  {name:'マーベラス',ticker:'7844',category:'video'},
  // 音楽
  {name:'エイベックス',ticker:'7860',category:'music'},
  {name:'amuse',ticker:'4301',category:'music'},
  // ニュース
  {name:'セレンディップHD',ticker:'7318',category:'news'},
  // 広告テック
  {name:'フリークアウトHD',ticker:'6094',category:'adtech'},
  {name:'ログリー',ticker:'6579',category:'adtech'},
  {name:'Macbee Planet',ticker:'7095',category:'adtech'},
  {name:'CARTA HD',ticker:'3688',category:'adtech'},
  {name:'SMN',ticker:'6185',category:'adtech'},
  {name:'Speee',ticker:'4499',category:'adtech'},
  {name:'バリューコマース',ticker:'2491',category:'adtech'},
  {name:'FRONTEO',ticker:'2158',category:'adtech'},
  {name:'ホットリンク',ticker:'3680',category:'adtech'},
  {name:'オリコン',ticker:'4800',category:'adtech'},
  {name:'イード',ticker:'6038',category:'adtech'},
  {name:'セレス',ticker:'3696',category:'adtech'},
  {name:'レントラックス',ticker:'6045',category:'adtech'},
  {name:'イトクロ',ticker:'6049',category:'adtech'},
  {name:'ネットイヤーグループ',ticker:'3622',category:'adtech'},
  {name:'アステリア',ticker:'3853',category:'adtech'},
  // 印刷・DX
  {name:'大日本印刷',ticker:'7912',category:'print_dx'},
  {name:'凸版印刷',ticker:'7911',category:'print_dx'},
  {name:'共同印刷',ticker:'7914',category:'print_dx'},
  // 追加メディア関連
  {name:'ブシロード',ticker:'7803',category:'video'},
  {name:'ブロッコリー',ticker:'2706',category:'publishing'},
  {name:'ボルテージ',ticker:'3639',category:'publishing'},
  {name:'ディー・エル・イー',ticker:'3686',category:'video'},
  {name:'ケイブ',ticker:'3760',category:'video'},
  {name:'シリコンスタジオ',ticker:'3907',category:'video'},
  {name:'タカラトミー',ticker:'7867',category:'publishing'},
  {name:'フィールズ',ticker:'2767',category:'video'},
  {name:'テイクアンドギヴ・ニーズ',ticker:'4331',category:'video'},
  {name:'sMedio',ticker:'3913',category:'adtech'},
  {name:'ピクセルカンパニーズ',ticker:'2743',category:'adtech'},
  {name:'サン電子',ticker:'6736',category:'adtech'},
];

const BATCH=5, DELAY=300;
const sleep=ms=>new Promise(r=>setTimeout(r,ms));

async function fetchOne(ticker){
  const sym=ticker+'.T';
  try{
    const [q,s]=await Promise.all([
      yf.quote(sym),
      yf.quoteSummary(sym,{modules:['financialData','defaultKeyStatistics']}).catch(()=>null)
    ]);
    let opm=null,roe=null,per=null,pbr=null;
    if(s){const fd=s.financialData||{},ks=s.defaultKeyStatistics||{};
      opm=fd.operatingMargins!=null?+(fd.operatingMargins*100).toFixed(1):null;
      roe=fd.returnOnEquity!=null?+(fd.returnOnEquity*100).toFixed(1):null;
      per=ks.trailingPE??ks.forwardPE??null;pbr=ks.priceToBook??null;
      if(per!=null)per=+per.toFixed(1);if(pbr!=null)pbr=+pbr.toFixed(2);}
    if(!per&&q.trailingPE)per=+q.trailingPE.toFixed(1);
    if(!pbr&&q.priceToBook)pbr=+q.priceToBook.toFixed(2);
    return{price:q.regularMarketPrice??null,marketCap:q.marketCap?Math.round(q.marketCap/1e8):null,operatingMargin:opm,roe,per,pbr};
  }catch(e){return null;}
}

async function main(){
  console.log('=== デジタルメディア: '+COMPANIES.length+'社 ===');
  const results=[];
  for(let i=0;i<COMPANIES.length;i+=BATCH){
    const batch=COMPANIES.slice(i,i+BATCH);
    const br=await Promise.all(batch.map(c=>fetchOne(c.ticker).then(d=>({...c,data:d}))));
    br.forEach(r=>{if(r.data){results.push({name:r.name,ticker:r.ticker,category:r.category,...r.data});process.stdout.write('.');}else process.stdout.write('x');});
    if(i+BATCH<COMPANIES.length)await sleep(DELAY);
  }
  console.log('\n→ '+results.length+'/'+COMPANIES.length+'社');
  const out=path.join(__dirname,'..','data','digital-media-companies.json');
  fs.writeFileSync(out,JSON.stringify(results,null,2),'utf8');
  console.log('Saved: '+out);
}
main().catch(e=>{console.error(e);process.exit(1);});
