import axios from 'axios';
import PathBuilder from './../pathBuilder';
import * as utility from './../utility';

export default class File {
	constructor(parent, element) {
		this.name = 'file';
		this.pathBuilder = new PathBuilder;
		this.parent = parent;
		this.contextUrl = this.parent.contextUrl;
		this.clientContext = this.parent.clientContext;
		this.element = element;
		this.pathBuilder.set('contextUrl', this.contextUrl);
		this.isList = this.parent.name === 'list';
		this.pathBuilder.set('element', this.parent.element);
		this.executeBinded = this.execute.bind(this, this.name, this.element);
	}

	async get(opts = {}) {
		if (opts.blob) {
			return this.executeBinded(async element => {
				return {
					request: {
						url: `${this._getSPObject(element, true)}/$value`,
						binaryStringResponseBody: true
					}
				}
			}, {
					rest: true,
					getter: res => res.body
				})
		} else {
			return this.executeBinded(async element => this._getSPObject(element, opts.rest), opts)
		}
	}
	async create(opts = {}) {
		opts.actionType = 'create';
		opts.logTextToAdd = this.parent.element ? ` in ${this.parent.element}` : '';
		if (opts.rest) {
			return this.executeBinded(async element => {
				if (typeOf(element.Content) === 'file') {
					return this._createWithRESTFromFile(element)
				} else {
					return this._createWithREST(element)
				}
			}, opts)
		} else {
			return this.executeBinded(async element => {
				return this._createWithJSOM(element)
			}, opts)
		}
	}
	async _createWithRESTFromFile(element) {
		let founds;
		let inputs = [];
		let {
			Url,
			Content,
			Overwrite,
			OnProgress = () => { }
		} = element;
		let {
			listName,
			filename,
			folder
		} = this._parseElementUrl(Url);
		if (!listName) {
			throw new Error('File creation url does not belong to list');
			return;
		}
		let requiredInputs = {
			__REQUESTDIGEST: true,
			__VIEWSTATE: true,
			__EVENTTARGET: true,
			__EVENTVALIDATION: true,
			ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation: true,
			ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle: true,
		}
		let listGUID = (await this.parent.get_('Id', {
			cached: true
		})).toString();
		let res = await axios.get(`${this.contextUrl}/_layouts/15/Upload.aspx?List={${listGUID}}`);
		let formMatches = res.data.match(/<form(\w|\W)*<\/form>/);
		let inputRE = /<input[^<]*\/>/g;
		while (founds = inputRE.exec(formMatches)) {
			let item = founds[0];
			let id = item.match(/id=\"([^\"]+)\"/)[1];
			if (requiredInputs[id]) {
				if (id === '__EVENTTARGET') item = item.replace(/value="[^\"]*"/, 'value="ctl00$PlaceHolderMain$ctl03$RptControls$btnOK"');
				if (id === 'ctl00_PlaceHolderMain_ctl04_ctl01_uploadLocation' && folder) item = item.replace(/value="[^\"]*"/, `value="/${folder}"`);
				if (id === 'ctl00_PlaceHolderMain_UploadDocumentSection_ctl05_OverwriteSingle' && Overwrite !== true) item = item.replace(/checked="[^\"]*"/, '');
				inputs.push(item);
			}
		}
		let form = window.document.createElement('form');
		form.innerHTML = inputs.join('');
		let formData = new FormData(form);
		formData.append('ctl00$PlaceHolderMain$UploadDocumentSection$ctl05$InputFile', Content, filename);
		return {
			request: {
				url: `${this.contextUrl}/_layouts/15/UploadEx.aspx?List={${listGUID}}`,
				method: 'POST',
				data: formData,
				onUploadProgress: e => OnProgress(Math.floor((e.loaded * 100) / e.total))
			},
			params: {
				errorHandler: res => {
					if (/ctl00_PlaceHolderMain_LabelMessage/i.test(res.data)) throw new Error(res.data.match(/ctl00_PlaceHolderMain_LabelMessage">([^<]+)<\/span>/)[1])
				},
				httpProvider: axios
			}
		}
	}
	async _createWithREST(element) {
		let {
			Url,
			Content,
			Overwrite
		} = element;
		let {
			listName,
			filename,
			folder
		} = this._parseElementUrl(Url);
		if (folder) await this.parent.folder(folder).create({
			silent: true,
			raw: true,
			view: ['Name'],
			cached: true
		});
		return {
			request: {
				url: `${this._getSPObject(folder, true)}/add(url='${filename}',overwrite=${!!Overwrite})`,
				method: 'POST',
				body: Content
			}
		}
	}
	async _createWithJSOM(element) {
		let {
			Url,
			Content,
			Overwrite
		} = element;
		let {
			filename,
			folder
		} = this._parseElementUrl(Url);
		folder && await this.parent.folder(folder).create({
			silent: true,
			raw: true,
			view: ['Name'],
			cached: true
		});
		let createInfo = new SP.FileCreationInformation;
		createInfo.set_url(filename);
		createInfo.set_content(this.convertFileContent(Content));
		createInfo.set_overwrite(!!Overwrite);
		return this._getSPObject(folder).add(createInfo)
	}
	async update(opts = {}) {
		return this.executeBinded(async element => {
			let {
				Content,
				columns
			} = element;
			let file = this._getSPObject(element);
			for (let columnName in columns) {
				let column = columns[columnName];
				if (typeOf(column) === 'array') columns[columnName] = column.join(';#;#');
			}
			let binaryInfo = new SP.FileSaveBinaryInformation;
			if (Content !== void 0) binaryInfo.set_content(this.convertFileContent(Content));
			binaryInfo.set_fieldValues(columns);
			file.saveBinary(binaryInfo);
			return file
		}, {
				...opts,
				actionType: 'update',
				logTextToAdd: this.parent.element ? ` in ${this.parent.element}` : ''
			});
	}
	async delete(opts = {}) {
		let {
			noRecycle,
			mask,
		} = opts;
		let delMethod = noRecycle ? 'deleteObject' : 'recycle';
		opts.logTextToAdd = this.parent.element ? ` in ${this.parent.element}` : '';
		opts.actionType = delMethod === 'recycle' ? 'recycle' : 'delete';
		return this.executeBinded(async element => {
			let spObject = this._getSPObject(element);
			if (spObject.getEnumerator) {
				let files = await this.parent.file(element).get({
					view: ['ServerRelativeUrl'],
					expanded: true
				});
				await this.execute(this.name, files, spElement => {
					let serverRelativeUrl = spElement.get_serverRelativeUrl();
					if (!mask || (mask && mask.test(serverRelativeUrl))) {
						spElement.cacheUrl = serverRelativeUrl;
						spElement[delMethod]();
						return spElement;
					}
				}, opts)
			} else {
				spObject[delMethod]()
				return spObject;
			}
		}, opts)
	}
	async copy(to, opts = {}) {
		return this._moveOrCopy(to, {
			...opts,
			actionType: 'copy'
		});
	}
	async move(to, opts = {}) {
		return this._moveOrCopy(to, {
			...opts,
			actionType: 'move',
			isMove: true
		});
	}
	async _moveOrCopy(to, opts) {
		let {
			mask,
			isMove
		} = opts;
		let method = isMove ? 'moveTo' : 'copyTo';
		let pathBuilder = new PathBuilder;
		pathBuilder.set('parentElementUrl', this.parent.element);
		pathBuilder.set('destName', to);
		return this.executeBinded(async element => {
			let pathBuilder = new PathBuilder;
			pathBuilder.set('parentElementUrl', this.parent.element);
			pathBuilder.set('destName', to);
			let spObject = this._getSPObject(element);
			if (spObject.getEnumerator) {
				let files = await this.parent.file(element).get({
					view: ['ServerRelativeUrl', 'Name'],
					expanded: true
				});
				await this.execute(this.name, files, spElement => {
					let serverRelativeUrl = spElement.get_serverRelativeUrl();
					if (!mask || (mask && mask.test(serverRelativeUrl))) {
						if (!/\./.test(pathBuilder.get('destName'))) pathBuilder.set('filename', spElement.get_name());
						spElement.cacheUrl = serverRelativeUrl;
						spElement[method](pathBuilder.build());
						return spElement;
					}
				}, opts)
			} else {
				spObject[method](pathBuilder.build());
			}
			return spObject;
		}, {
				...opts,
				logTextToAdd: ` -> ${pathBuilder.build()}`
			});
	}

	_parseElementUrl(element) {
		let folder, filename;
		if (element) {
			let fileSplits = element.split('/');
			if (this.isList && fileSplits[0] === this.parent.element) fileSplits.shift();
			if (/\./.test(fileSplits[fileSplits.length - 1])) filename = fileSplits.pop();
			folder = fileSplits.join('/');
		}
		return {
			filename,
			folder
		}
	}

	_getSPObject(element, isREST) {
		if (typeOf(element) === 'object') element = element.Url;
		let {
			filename,
			folder
		} = this._parseElementUrl(element);
		this.pathBuilder.set('folder', folder);
		this.pathBuilder.set('filename', filename);
		let fullElementUrl = this.pathBuilder.build();
		if (isREST) {
			let url = `${this.contextUrl}/_api/web`;
			if (filename) {
				return `${url}/getfilebyserverrelativeurl('/${fullElementUrl}')`
			} else {
				return `${url}/getfolderbyserverrelativeurl('/${fullElementUrl}')/files`
			}
		} else {
			let web = this.clientContext.get_web();
			if (fullElementUrl) {
				if (filename) {
					return web.getFileByServerRelativeUrl(`/${fullElementUrl}`)
				} else {
					return web.getFolderByServerRelativeUrl(`/${fullElementUrl}`).get_files();
				}
			} else {
				return web.get_rootFolder().get_files();
			}
		}
	}
}