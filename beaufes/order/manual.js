// 版表示はこの定数だけで管理する。
const GUIDE_VERSION='v1.1.1';
document.getElementById('guide-version').textContent=GUIDE_VERSION;
document.getElementById('print-guide').addEventListener('click',()=>window.print());
