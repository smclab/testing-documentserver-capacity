const crypto = require('crypto');
const request = require('request');
const fs = require('fs');
const W3CWebSocket = require('websocket').w3cwebsocket;
const extraEscapable = /[\x00-\x1f\ud800-\udfff\ufffe\uffff\u0300-\u0333\u033d-\u0346\u034a-\u034c\u0350-\u0352\u0357-\u0358\u035c-\u0362\u0374\u037e\u0387\u0591-\u05af\u05c4\u0610-\u0617\u0653-\u0654\u0657-\u065b\u065d-\u065e\u06df-\u06e2\u06eb-\u06ec\u0730\u0732-\u0733\u0735-\u0736\u073a\u073d\u073f-\u0741\u0743\u0745\u0747\u07eb-\u07f1\u0951\u0958-\u095f\u09dc-\u09dd\u09df\u0a33\u0a36\u0a59-\u0a5b\u0a5e\u0b5c-\u0b5d\u0e38-\u0e39\u0f43\u0f4d\u0f52\u0f57\u0f5c\u0f69\u0f72-\u0f76\u0f78\u0f80-\u0f83\u0f93\u0f9d\u0fa2\u0fa7\u0fac\u0fb9\u1939-\u193a\u1a17\u1b6b\u1cda-\u1cdb\u1dc0-\u1dcf\u1dfc\u1dfe\u1f71\u1f73\u1f75\u1f77\u1f79\u1f7b\u1f7d\u1fbb\u1fbe\u1fc9\u1fcb\u1fd3\u1fdb\u1fe3\u1feb\u1fee-\u1fef\u1ff9\u1ffb\u1ffd\u2000-\u2001\u20d0-\u20d1\u20d4-\u20d7\u20e7-\u20e9\u2126\u212a-\u212b\u2329-\u232a\u2adc\u302b-\u302c\uaab2-\uaab3\uf900-\ufa0d\ufa10\ufa12\ufa15-\ufa1e\ufa20\ufa22\ufa25-\ufa26\ufa2a-\ufa2d\ufa30-\ufa6d\ufa70-\ufad9\ufb1d\ufb1f\ufb2a-\ufb36\ufb38-\ufb3c\ufb3e\ufb40-\ufb41\ufb43-\ufb44\ufb46-\ufb4e\ufff0-\uffff]/g;
var extraLookup;

// This may be quite slow, so let's delay until user actually uses bad
// characters.
var unrollLookup = function (escapable) {
	var i;
	var unrolled = {};
	var c = [];
	for (i = 0; i < 65536; i++) {
		c.push(String.fromCharCode(i));
	}
	escapable.lastIndex = 0;
	c.join('').replace(escapable, function (a) {
		unrolled[a] = '\\u' + ('0000' + a.charCodeAt(0).toString(16)).slice(-4);
		return '';
	});
	escapable.lastIndex = 0;
	return unrolled;
};
function quote(string) {
	var quoted = JSON.stringify(string);

	// In most cases this should be very fast and good enough.
	extraEscapable.lastIndex = 0;
	if (!extraEscapable.test(quoted)) {
		return quoted;
	}

	if (!extraLookup) {
		extraLookup = unrollLookup(extraEscapable);
	}

	return quoted.replace(extraEscapable, function (a) {
		return extraLookup[a];
	});
}
function randomString(count = 8) {
	return crypto.randomBytes(count).toString('hex');
}
function randomNumber(count = 1000) {
	return Math.floor(Math.random() * count) + '';
}

const enableLog = -1 !== process.argv.indexOf('--debug');
function log(docId, message) {
	enableLog && console.log(new Date().toLocaleString() + ' ' + docId + ': ' + message);
}

function DocsCoApi(options = {}) {
	this.docId = options.docId || '1234567890';
	this.server = options.server || 'ws://127.0.0.1:8001';
	this.url = options.url || 'https://doc.onlyoffice.com/example/samples/sample.docx';
	this.sessionId = randomString();
	this.serverId = randomNumber();
	this.client = null;
	this.uid = null;
	this.isWord = false;
	this.changeCount = 100;
	this.init();
}
DocsCoApi.prototype.init = function () {
	this.client =
		new W3CWebSocket(this.server + '/doc/' + this.docId + '/c/' + this.serverId + '/' + this.sessionId + '/websocket');
	this.client.onerror = () => {
		this._log('Connection Error');
	};

	this.client.onopen = () => {
		this._log('WebSocket Client Connected');
	};

	this.client.onclose = () => {
		this._log('echo-protocol Client Closed');
	};

	this.client.onmessage = (e) => {
		const msg = e.data;
		if (typeof msg === 'string') {
			this._log('Received: "' + msg + '"');

			const type = msg.slice(0, 1);
			const content = msg.slice(1);
			var payload;

			if (content) {
				try {
					payload = JSON.parse(content);
				} catch (e) {
					this._log('bad json', content);
				}
			}

			switch (type) {
				case 'o':
					this._log('open');
					break;
				case 'h':
					this._log('heartbeat');
					break;
				case 'a':
					if (Array.isArray(payload)) {
						payload.forEach((p) => {
							this.onMessage(p);
						});
					}
					break;
				case 'm':
					this.onMessage(payload);
					break;
				case 'c':
					this._log('close');
					break;
			}
		}
	};
};
DocsCoApi.prototype._log = function (message) {
	log(this.docId, message);
};
DocsCoApi.prototype._onAuth = function (data) {
	this._log('onAuth ');
	console.log(data);

	var participants = data['participants'];
	//console.log(participants[0].id);
	this.uid = participants[0].id;
};
DocsCoApi.prototype._onMessages = function (data) {
	this._log('onMessages: ' + data["messages"]);
	console.log(data);

	this.doWordChange();
};
DocsCoApi.prototype._onCursor = function (data) {
	this._log('onCursor ');
	console.log(data);
};
DocsCoApi.prototype._onGetLock = function (data) {
	this._log('onGetLock ');
	console.log(data);
};
DocsCoApi.prototype._onReleaseLock = function (data) {
	this._log('onReleaseLock ');
	console.log(data);
};
DocsCoApi.prototype._onConnectionStateChanged = function (data) {
	this._log('onConnectionStateChanged ');
	console.log(data);
};
DocsCoApi.prototype._onSaveChanges = function (data) {
	this._log('onSaveChanges ');
	console.log(data);
};
DocsCoApi.prototype._onSaveLock = function (data) {
	this._log('onSaveLock ');
	console.log(data);
};
DocsCoApi.prototype._onUnSaveLock = function (data) {
	this._log('onUnSaveLock ');
	console.log(data);

	this.doWordChange();
};
DocsCoApi.prototype._onSavePartChanges = function (data) {
	this._log('onSavePartChanges ');
	console.log(data);
};
DocsCoApi.prototype._onDrop = function (data) {
	this._log('onDrop ');
	console.log(data);
};
DocsCoApi.prototype._documentOpen = function (data1) {
	this.sendRequest({'type': 'getMessages'});
	this._log('documentOpen ');
	var data = data1.toString();
	// Removed: Fix broken url
	console.log(data);

	if ((data = data['data']) && (data = data['data']) && (data = data['Editor.bin'])) {
		request(data).pipe(fs.createWriteStream('files/' + randomString() + '-' + 'Editor.bin'));
		return;
	}

	this._log('error open file: ' + this.url);
};
DocsCoApi.prototype._onWarning = function (data) {
	this._log('onWarning ');
	console.log(data);
};
DocsCoApi.prototype._onLicense = function (data) {
	this.sendRequest({
		'type': 'auth',
		'docid': this.docId,
		'token': 'fghhfgsjdgfjs',
		'user': {'id': 'uid-1', 'username': 'Jonn Smith', 'indexUser': -1},
		'editorType': 1,
		'lastOtherSaveTime': -1,
		'block': [],
		'sessionId': null,
		'view': false,
		'isCloseCoAuthoring': false,
		'openCmd': {
			'c': 'open',
			'id': this.docId,
			'userid': 'uid-1',
			'format': 'docx',
			'url': this.url,
			'title': 'test',
			'embeddedfonts': false,
			'viewmode': false
		},
		'version': '3.0.9'
	});
};
DocsCoApi.prototype.onMessage = function (data) {
	this._log('message: "' + data + '"');
	console.log(data);

	var dataObject = JSON.parse(data);
	switch (dataObject['type']) {
		case 'auth'        :
			this._onAuth(dataObject);
			break;
		case 'message'      :
			this._onMessages(dataObject, false);
			break;
		case 'cursor'       :
			this._onCursor(dataObject);
			break;
		case 'getLock'      :
			this._onGetLock(dataObject);
			break;
		case 'releaseLock'    :
			this._onReleaseLock(dataObject);
			break;
		case 'connectState'    :
			this._onConnectionStateChanged(dataObject);
			break;
		case 'saveChanges'    :
			this._onSaveChanges(dataObject);
			break;
		case 'saveLock'      :
			this._onSaveLock(dataObject);
			break;
		case 'unSaveLock'    :
			this._onUnSaveLock(dataObject);
			break;
		case 'savePartChanges'  :
			this._onSavePartChanges(dataObject);
			break;
		case 'drop'        :
			this._onDrop(dataObject);
			break;
		case 'waitAuth'      : /*Ждем, когда придет auth, документ залочен*/
			break;
		case 'error'      : /*Старая версия sdk*/
			this._onDrop(dataObject);
			break;
		case 'documentOpen'    :
			this._documentOpen(dataObject);
			break;
		case 'warning':
			this._onWarning(dataObject);
			break;
		case 'license':
			this._onLicense(dataObject);
			break;
	}
};
DocsCoApi.prototype.sendRequest = function (data) {
	if (this.client.readyState === this.client.OPEN) {
		const sendData = JSON.stringify(data);
		this._log("Send: '" + sendData + "'");
		this.client.send(quote(sendData));
	}
};
DocsCoApi.prototype.doWordChange = function () {
	if (this.changeCount == 0) {
		this.sendRequest({'type': 'close'});
		//this.client.close();
		return;
	}

	this._log('do change #' + this.changeCount--);

	this.sendRequest({
		type: 'saveChanges',
  		changes: '["80;AgAAADEAAQAAAP//AAAdcq7bTWYAAC0BAAAEAAAAAAAAAAAAAAABAAAAAAAAAPb///8aAAAANAAuADEALgA1AC4AMQAuAEAAQABSAGUAdgA=","35;BgAAADEANQAwABwAAAABAAAAAQEAAAAAAAAAAQAAAHQAAAA=","35;BgAAADEANQAwABwAAAABAAAAAQEAAAABAAAAAQAAAGUAAAA=","35;BgAAADEANQAwABwAAAABAAAAAQEAAAACAAAAAQAAAHMAAAA=","35;BgAAADEANQAwABwAAAABAAAAAQEAAAADAAAAAQAAAHQAAAA=","80;AgAAADEAAQAAAP//AAAdcq7bTWYAAIsAAAABAAAAAQAAAAAAAAABAAAAAAAAAPb///8aAAAANAAuADEALgA1AC4AMQAuAEAAQABSAGUAdgA=","35;BgAAADEANQAwABwAAAABAAAAAQEAAAAEAAAAAgAAAAAAAAA="]',
  		startSaveChanges: true,
  		endSaveChanges: true,
  		isCoAuthoring: false,
  		isExcel: false,
  		deleteIndex: null,
  		excelAdditionalInfo: '{"Gk":"' + this.uid + '","B4c":"uid-1","OLc":"14;BgAAADEANQAwAAUAAAA="}'
	});
};

var countUsers = 1;
var countDocuments = 1;
var serverUrl, documentUrl, baseUrl;

var indexArg = process.argv.indexOf('--users');
if (-1 !== indexArg) {
	countUsers = process.argv[indexArg + 1];
}
indexArg = process.argv.indexOf('--documents');
if (-1 !== indexArg) {
	countDocuments = process.argv[indexArg + 1];
}
indexArg = process.argv.indexOf('--server');
if (-1 !== indexArg) {
	serverUrl = process.argv[indexArg + 1];
}
indexArg = process.argv.indexOf('--file');
if (-1 !== indexArg) {
	documentUrl = process.argv[indexArg + 1];
}
indexArg = process.argv.indexOf('--base-url');
if (-1 !== indexArg) {
	baseUrl = process.argv[indexArg + 1];
}

var docNames = [
//'test.odt', 'test.docx'
 'test.docx'
];
//'msoffice1.xlsx' ];

var now = new Date();

var prefix = dateFormat(now, 'mmddHHMM');

var sDocId;
for (var nUser = 0; nUser < countUsers; ++nUser) {
	for (var nDoc = 0; nDoc < countDocuments; ++nDoc) {
		sDocId = prefix + '_' + nUser + '_' + nDoc + '_' + randomString();

		if (baseUrl) {
			var x = (nDoc * nUser)  % docNames.length;

			documentUrl = baseUrl + '/' + docNames[x];
		}
		var oDocsCoApi = new DocsCoApi({server: serverUrl, docId: sDocId, url: documentUrl});
	}
}
