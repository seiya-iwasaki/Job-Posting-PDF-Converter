import './style.css';
import './app.css';
import { SaveXLSXsToPDFDir } from '../wailsjs/go/internal/App';

// ロゴなしのドラッグ&ドロップUI
window.addEventListener('DOMContentLoaded', () => {
  const app = document.querySelector('#app');
  if (!app) return;
  app.innerHTML = `
    <div class="drop-area" id="drop-area">
      <p>ここにファイルまたはフォルダをドラッグ＆ドロップしてください</p>
      <input type="file" id="fileElem" multiple webkitdirectory directory style="display:none" />
      <button class="btn" id="fileSelectBtn">ファイル/フォルダを選択</button>
    </div>
    <div id="error-list" style="color:#b85c3b; margin-bottom:10px; font-size:0.97em;"></div>
    <ul id="file-list"></ul>
    <button class="btn" id="sendBtn" style="margin-top:18px;">PDFに変換</button>
  `;

  const dropArea = document.getElementById('drop-area');
  const fileElem = document.getElementById('fileElem');
  const fileList = document.getElementById('file-list');
  const fileSelectBtn = document.getElementById('fileSelectBtn');
  const errorList = document.getElementById('error-list');
  const sendBtn = document.getElementById('sendBtn');
  let lastXlsxFiles = [];

  // ドラッグ時のスタイル変更
  ['dragenter', 'dragover'].forEach(eventName => {
    dropArea.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropArea.classList.add('highlight');
    }, false);
  });
  ['dragleave', 'drop'].forEach(eventName => {
    dropArea.addEventListener(eventName, (e) => {
      e.preventDefault();
      e.stopPropagation();
      dropArea.classList.remove('highlight');
    }, false);
  });

  sendBtn.disabled = true;

  dropArea.addEventListener('drop', handleDrop, false);
  fileSelectBtn.addEventListener('click', () => fileElem.click());
  fileElem.addEventListener('change', (e) => {
    handleFiles(e.target.files);
  });

  sendBtn.addEventListener('click', async () => {
    if (lastXlsxFiles.length === 0) {
      sendBtn.innerHTML = 'xlsxファイルが選択されていません';
      return;
    }
    // ファイルをBase64でまとめてGoに送信
    const fileDatas = await Promise.all(lastXlsxFiles.map(async (file) => {
      const data = await fileToBase64(file);
      return { name: file.name, data };
    }));
    try {
      await SaveXLSXsToPDFDir(fileDatas);
      sendBtn.innerHTML = '変換しました';
    } catch (e) {
      sendBtn.innerHTML = 'エラー: ' + e;
    }
  });

  function fileToBase64(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.onload = () => {
        // DataURL: "data:...;base64,..." → base64部分だけ抽出
        const base64 = reader.result.split(',')[1];
        resolve(base64);
      };
      reader.onerror = reject;
      reader.readAsDataURL(file);
    });
  }

  function handleDrop(e) {
    const dt = e.dataTransfer;
    errorList.innerHTML = '';
    if (dt.items) {
      // フォルダ対応
      const items = Array.from(dt.items);
      let entries = items.map(item => item.webkitGetAsEntry && item.webkitGetAsEntry());
      entries = entries.filter(Boolean);
      if (entries.length > 0) {
        fileList.innerHTML = '';
        errorList.innerHTML = '';
        lastXlsxFiles = [];
        entries.forEach(entry => traverseFileTree(entry));
        return;
      }
    }
    handleFiles(dt.files);
  }

  function handleFiles(files) {
    fileList.innerHTML = '';
    errorList.innerHTML = '';
    sendBtn.innerHTML = 'PDFに変換';
    const nonXlsxFiles = [];
    const xlsxFiles = [];
    Array.from(files).forEach(file => {
      if (file.name.toLowerCase().endsWith('.xlsx')) {
        xlsxFiles.push(file);
        const li = document.createElement('li');
        li.textContent = file.webkitRelativePath || file.name;
        fileList.appendChild(li);
      } else {
        nonXlsxFiles.push(file);
      }
    });
    lastXlsxFiles = xlsxFiles;
    if (nonXlsxFiles.length > 0) {
      sendBtn.disabled = true;
      const names = nonXlsxFiles.map(f => (f.webkitRelativePath || f.name)).join('<br>');
      errorList.style.display = 'block';
      errorList.innerHTML = 'xlsx以外のファイルが含まれています:<br>' + names;
    } else {
      sendBtn.disabled = false;
      errorList.style.display = 'none';
    }
  }

  // フォルダ内のファイルも再帰的に取得
  function traverseFileTree(item, path = "") {
    sendBtn.innerHTML = 'PDFに変換';
    sendBtn.disabled = false;
    if (item.isFile) {
      item.file(file => {
        if (file.name.toLowerCase().endsWith('.xlsx')) {
          const li = document.createElement('li');
          li.textContent = path + file.name;
          fileList.appendChild(li);
          lastXlsxFiles.push(file);
        } else {
          // 画面上にエラー表示
          sendBtn.disabled = true;
          const msg = path + file.name;
          if (!errorList.innerHTML.includes(msg)) {
            if (errorList.innerHTML === '' || errorList.style.display === 'none') {
              errorList.innerHTML = 'xlsx以外のファイルが含まれています:<br>';
              errorList.style.display = 'block';
            }
            errorList.innerHTML += msg + '<br>';
          }
        }
      });
    } else if (item.isDirectory) {
      const dirReader = item.createReader();
      dirReader.readEntries(entries => {
        entries.forEach(entry => {
          traverseFileTree(entry, path + item.name + "/");
        });
      });
    }
  }
});
