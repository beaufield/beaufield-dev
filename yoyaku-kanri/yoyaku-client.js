/* 予約DB通信。JWTはメモリのみ、保存前に操作UUIDを永続化し、結果不明を自動再送しない。 */
(function(root) {
  'use strict';
  class UnknownResult extends Error {
    constructor(requestId) { super('保存結果を確認しています。再登録せず、操作結果を確認してください。'); this.name='UnknownResult'; this.requestId=requestId; }
  }
  class YoyakuClient {
    constructor({url,publicKey,issueToken,session,storage,fetchImpl,uuid,now}) {
      if(!/^https:\/\/[a-z0-9]{20}\.supabase\.co$/.test(url))throw Error('接続先設定を確認してください');
      if(typeof publicKey!=='string'||!publicKey||publicKey.startsWith('sb_secret_'))throw Error('公開キーが必要です');
      if(publicKey.startsWith('eyJ')) { try { if(JSON.parse(atob(publicKey.split('.')[1])).role==='service_role')throw Error('PRIVATE_KEY_FORBIDDEN'); } catch(e) { throw Error('公開用publishableキーを使用してください'); } }
      this.url=url;this.publicKey=publicKey;this.issueToken=issueToken;this.session=session;
      this.storage=storage||root.localStorage;this.fetch=fetchImpl||root.fetch.bind(root);
      this.uuid=uuid||(()=>root.crypto.randomUUID());this.now=now||Date.now;
      this.jwt=null;this.jwtPromise=null;this.snapshot=null;this.scope=null;
    }
    identity() {
      const s=this.session();
      if(!s||!s.user_id||!s.token||!(s.expires>this.now()))throw Error('SESSION_INVALID');
      const key=`bf_yoyaku_db_v1:${new URL(this.url).hostname}:${encodeURIComponent(s.user_id)}`;
      if(this.scope!==key||this.sessionToken!==s.token) {this.scope=key;this.sessionToken=s.token;this.jwt=null;this.snapshot=null;}
      return {s,key};
    }
    async token() {
      const {s}=this.identity();
      if(this.jwt&&this.jwt.exp*1000>this.now()+30000)return this.jwt.token;
      if(!this.jwtPromise) {
        const issuedFor=s.token;
        this.jwtPromise=(async()=>{
          const r=await this.issueToken(s.token);
          if(!r||!r.success||!r.token||!Number.isFinite(r.exp))throw Error(r?.error||'AUTH_UNAVAILABLE');
          if(r.exp*1000<=this.now()||r.exp*1000>Math.min(s.expires,this.now()+301000))throw Error('INVALID_TOKEN_EXPIRY');
          if(this.session()?.token!==issuedFor)throw Error('SESSION_CHANGED');
          this.jwt={token:r.token,exp:r.exp};return r.token;
        })().finally(()=>{this.jwtPromise=null;});
      }
      return this.jwtPromise;
    }
    async rpc(name,args,readRetry=0) {
      this.identity();const originalScope=this.scope,originalToken=this.sessionToken;
      const token=await this.token();
      const controller=new AbortController();const timer=setTimeout(()=>controller.abort(),20000);
      let response;
      try {
        response=await this.fetch(`${this.url}/rest/v1/rpc/${name}`,{method:'POST',headers:{apikey:this.publicKey,Authorization:`Bearer ${token}`,'Content-Type':'application/json','Content-Profile':'yoyaku_api','Accept-Profile':'yoyaku_api'},body:JSON.stringify(args),signal:controller.signal});
        const result=await response.json();
        this.identity();if(this.scope!==originalScope||this.sessionToken!==originalToken)throw Error('SESSION_CHANGED');
        if(!response.ok) {
          const authFailure=response.status===401||/^YOYAKU_(AUTH_REQUIRED|ACCESS_DENIED)$/.test(result?.message||'');
          if(authFailure) {
            this.jwt=null;
            // 新しい役割の通行証を取り直すのは読取だけ。保存を自動再送しない。
            if(readRetry===0&&['yk_get_snapshot','yk_get_operation'].includes(name))return await this.rpc(name,args,1);
            if(name==='yk_get_snapshot')throw Error(response.status===401?'SESSION_INVALID':'ACCESS_DENIED');
          }
          throw Error('DB_HTTP_'+response.status);
        }
        return result;
      } finally {clearTimeout(timer);}
    }
    async readSnapshot() {
      this.identity();const requestScope=this.scope;
      const r=await this.rpc('yk_get_snapshot',{});
      if(!r?.success||!r.data||!Array.isArray(r.data.reservations)||r.data.reservations.length!==r.data.reservation_count)throw Error('一覧の全件取得を確認できません');
      this.identity();if(requestScope!==this.scope)throw Error('SESSION_CHANGED');
      const d=r.data;
      if(!Number.isSafeInteger(d.commit_seq)||!d.writer_epoch)throw Error('INVALID_SNAPSHOT');
      const sameEpoch=this.snapshot&&d.writer_epoch===this.snapshot.writer_epoch;
      if(!this.snapshot||(sameEpoch&&d.commit_seq>=this.snapshot.commit_seq)||(!sameEpoch&&Date.parse(d.server_time)>=Date.parse(this.snapshot.server_time))) {
        this.snapshot=d;
        try {this.storage.setItem(this.scope+':snapshot:'+d.writer_epoch,JSON.stringify(d));}catch(e) { /* 参照キャッシュの保存失敗は読取を妨げない。 */ }
      }
      return this.snapshot;
    }
    pending() {
      const {key}=this.identity();const raw=this.storage.getItem(key+':outbox');
      if(!raw)return null;
      let p;try {p=JSON.parse(raw);}catch(e) {throw Error('未確認操作の記録が壊れています。保存を停止しました');}
      if(!p||!p.request_id||!p.args||p.args.p_request_id!==p.request_id)throw Error('未確認操作の記録を確認してください');
      return p;
    }
    persist(p) {
      const raw=JSON.stringify(p),key=this.scope+':outbox';
      try {this.storage.setItem(key,raw);if(this.storage.getItem(key)!==raw)throw Error('READBACK_FAILED');}
      catch(e) {throw Error('端末に操作記録を保存できません。送信を中止しました');}
    }
    clearPending(id) {
      const p=this.pending();if(p?.request_id===id)this.storage.removeItem(this.scope+':outbox');
    }
    versions(action,payload) {
      const versions={};
      if(payload.product_id&&['saveProduct','toggleProductActive'].includes(action)) {
        const p=this.snapshot.products.find(p=>p.product_id===payload.product_id);
        if(!p)throw Error('商品を再読込してください');versions['product:'+p.product_id]=p.row_version;
      }
      const nos=payload.allocations?.map(a=>a.reservation_no)||payload.reservation_nos||(payload.reservation_no?[payload.reservation_no]:[]);
      for(const no of nos) {const r=this.snapshot.reservations.find(r=>r.reservation_no===Number(no));if(r)versions[String(no)]=r.row_version;}
      return versions;
    }
    accept(r) {
      if(!r?.success||!this.snapshot||r.writer_epoch!==this.snapshot.writer_epoch)return;
      if(r.commit_seq<=this.snapshot.commit_seq)return; // 古い返事で削除済み行を復活させない。
      if(r.commit_seq!==this.snapshot.commit_seq+1) {this.needsRefresh=true;return;}
      const d=r.delta;
      if(!d||!Array.isArray(d.reservations)||!Array.isArray(d.products)||!Array.isArray(d.removed_nos))throw Error('INVALID_DELTA');
      const rows=new Map(this.snapshot.reservations.map(x=>[x.reservation_no,x]));
      for(const no of d.removed_nos)rows.delete(no);
      for(const row of d.reservations)rows.set(row.reservation_no,row);
      const ps=new Map(this.snapshot.products.map(x=>[x.product_id,x]));for(const p of d.products)ps.set(p.product_id,p);
      this.snapshot={...this.snapshot,products:[...ps.values()],reservations:[...rows.values()].sort((a,b)=>a.reservation_no-b.reservation_no),reservation_count:rows.size,commit_seq:r.commit_seq};
    }
    async mutate(action,payload,displayed=null) {
      this.identity();if(this.pending())throw Error('未確認の保存があります。操作結果を先に確認してください');
      if(!this.snapshot||this.needsRefresh) {await this.readSnapshot();this.needsRefresh=false;}
      if(this.snapshot.writer!=='db')throw Error('現在は参照のみです');
      // 通行証取得の失敗を「DBへ送信済み」と混同しない。保存操作はまだ送らない。
      await this.token();
      const id=this.uuid();const args={p_request_id:id,p_epoch:displayed?.epoch||this.snapshot.writer_epoch,p_action:action,p_versions:displayed?.versions||this.versions(action,payload),p_payload:payload};
      this.persist({request_id:id,args,created_at:this.now(),state:'PREPARED'});
      let r;
      try {r=await this.rpc('yk_mutate',args);if(typeof r?.success!=='boolean'||r.request_id!==id)throw Error('INVALID_RESPONSE');}
      catch(e) {throw new UnknownResult(id);}
      this.accept(r);this.clearPending(id);return r;
    }
    async resolvePending() {
      const p=this.pending();if(!p)return {state:'NONE'};
      const r=await this.rpc('yk_get_operation',{p_request_id:p.request_id});
      if(['COMMITTED','REJECTED'].includes(r?.state)&&r.request_id===p.request_id) {
        this.accept(r.result);this.clearPending(p.request_id);return r;
      }
      // 見つからない≠未実行。新UUIDや自動再送はしない。
      return {state:'UNKNOWN',request_id:p.request_id};
    }
    async retryPendingExplicitly() {
      const p=this.pending();if(!p)throw Error('再確認する操作がありません');
      // UIの明示操作専用。同じUUID/epoch/payloadを保ち、内容を変更しない。
      let r;try {r=await this.rpc('yk_mutate',p.args);}catch(e) {throw new UnknownResult(p.request_id);}
      if(typeof r?.success!=='boolean'||r.request_id!==p.request_id)throw new UnknownResult(p.request_id);
      this.accept(r);this.clearPending(p.request_id);return r;
    }
  }
  root.YoyakuDb={YoyakuClient,UnknownResult};
})(typeof window==='undefined'?globalThis:window);
