export default class Searcher {
	constructor() {
		let it = this;
		let lookupTimeout;
		this.opts = opts;
		this.uri = new AuraURI;
		// console.log(this.uri);
		if (opts) {
			this.validTimeStamp;
			this.reqTimeStamp;
			this.listData = listData;
			this.spx = new SPX;
			this.folder = opts.folder;
			this.searchKey = opts.searchKey;
			this.listType = opts.listType;
			this.uri = opts.uriProvider || new AuraURI;
			this.value = opts.value || (this.searchKey && this.uri.query[this.searchKey] ? this.uri.query[this.searchKey] : '');
			this.columns = opts.columns || 'Title';
			this.orderBy = opts.orderBy || 'ID';
			this.orderByDate = opts.orderByDate;
			this.ascending = opts.ascending || false;
			this.pageSize = opts.pageSize || 5;
			this.recursive = opts.recursive;
			this.$inputContainer = (typeOf(opts.inputContainer) == 'string' ? $(opts.inputContainer) : opts.inputContainer);
			this.$outputContainer = (typeOf(opts.outputContainer) == 'string' ? $(opts.outputContainer) : opts.outputContainer);
			this.cb = opts.cb || function() {};
			this.preRender = opts.preRender;
			this.render = opts.render;
			this.postRender = opts.postRender;
			this.debugMode = opts.debugMode;
			if (opts.caml) {
				switch (typeOf(opts.caml)) {
					case 'object':
						this.caml = this.spx.getCamlQuery(opts.caml);
						break;
					case 'string':
						this.caml = this.spx.parseCamlString(opts.caml);
						break;
				}
			}
		}

		this.$outputContainer.html('\
			<div class="searcher-output-container"></div>\
			<div class="g-button-list js-button-loading" data-margin="1" data-state="2" data-pointer="1">\
				<div class="g-button-list_text">Показать еще</div>\
			</div>');
		this.$output = this.$outputContainer.find('.searcher-output-container');
		this.$button = this.$outputContainer.find('.js-button-loading');
		this.$buttonText = this.$button.find('.g-button-list_text');

		if (it.$inputContainer && it.$outputContainer) {
			if (listData) {
				this.spx.getListAllColumns(listData, this.makeForm.bind(this), {
					mapBy: 'InternalName',
					viewFields: ['Title', 'InternalName', 'TypeAsString'],
					isCached: true
				});
			} else {
				this.makeForm();
			}
		} else {
			console.log('inputContainer or outputContainer does not exist');
		}

		let prevValue = it.value;
	}

	makeForm(columnsData) {
		let lastValue = it.value;
		it.columnsData = columnsData;
		!it.value && it.searchFactory(it.value, it.fillPage);
		it.auraForm = new AuraForm([{
			labeled: {
				value: it.value,
				label: 'none',
				buttonIcon: 'magnifier',
				placeholder: 'Поиск',
				maxLength: 250,
				required: false,
				isValidateOnInput: true,
				isValidateOnBlur: false,
				isValidateOnPaste: true,
				validator: function(value, cb) {
					if (value) {
						it.uri.replaceQueryKey(it.searchKey, value, true);
					} else {
						it.uri.deleteQueryKey(it.searchKey, value, true);
					}
					clearTimeout(lookupTimeout);
					lookupTimeout = setTimeout(function() {
						lastValue = value;
						it.validTimeStamp = (new Date).getTime();
						it.reqTimeStamp = 0;
						it.$output.html('');
						it.searchFactory(value, function(results) {
							// console.log(results);
							let html = '';
							if (it.reqTimeStamp > it.validTimeStamp) {
								it.reqTimeStamp = (new Date).getTime();
								return;
							};
							if (results) {
								if (results && results.length === it.pageSize) {
									it.$buttonText.text('Показать еще');
									it.$button.off('click').click(function(e) {
										it._pager.moveNext();
									});
								} else {
									it.$button.text('Ничего нет');
									it.$button.off('click');
								}
								let deferred = $.Deferred();
								if (it.preRender) {
									it.preRender(results, deferred);
									$.when(deferred).done(function() {
										if (it.render) {
											it.renderItems(results)
										}
										it.cb(results);
										it.postRender && it.postRender(results);
										it.reqTimeStamp = (new Date).getTime();
										cb(true);
									})
								} else {
									if (it.render) {
										it.renderItems(results)
									}
									it.cb(results);
									it.postRender && it.postRender(results);
									it.reqTimeStamp = (new Date).getTime();
									cb(true);
								}
							} else {
								it.$button.text('Ничего нет');
								it.$button.off('click');
								cb(false);
							}
						})
					}, lastValue === value ? 0 : 250)
				}
			}
		}], null, {
			$: it.$inputContainer,
			submittable: false
		})
	}

	searchFactory(value, cb) {
		let it = this;
		this.reqTimeStamp = 0;
		this.validTimeStamp = Infinity;
		if (this.listData) {
			this._pager && this._pager.destroy && this._pager.destroy()
			this._pager = new Pager(it.listData, {
				isScrollable: true,
				folder: it.folder,
				listType: it.listType,
				orderBy: it.orderBy,
				orderByDate: it.orderByDate,
				caml: it.concatCaml(it.camlBuilder(value)),
				pageSize: it.pageSize,
				recursive: it.recursive,
				ascending: it.ascending,
				debugMode: it.debugMode,
				isCached: true,
				preload: function() {
					it.$button.text('Загрузка');
				},
				cb: function(res) {
					it.$button.text('Дальше');
					cb && cb.call(it, res)
				}
			});
		} else {
			if (value) {
				this.spx.searchItems(value, cb.bind(this), this.opts);
			} else {
				it.$button.text('Ничего нет');
			}
		}
	}

	fillPage(results) {
		// console.log(results.length);
		let it = this;
		let html = '';
		if (results) {
			if (results.length === it.pageSize) {
				it.$buttonText.text('Показать еще');
				it.$button.off('click').click(function(e) {
					it._pager.moveNext();
				})
			} else {
				it.$button.text('Ничего нет');
				it.$button.off('click');
			}
			let deferred = $.Deferred();
			if (it.preRender) {
				it.preRender(results, deferred);
				$.when(deferred).done(function() {
					let htmlItem;
					if (it.render) {
						it.renderItems(results)
					}
					it.cb(results);
					it.postRender && it.postRender(results)
				})
			} else {
				if (it.render) {
					it.renderItems(results)
				}
				it.cb(results);
				it.postRender && it.postRender(results)
			}
		} else {
			it.$button.text('Ничего нет');
			it.$button.off('click');
		}
	}

	renderItems(items) {
		let htmlItem;
		let html = '';
		for (let i = 0; i < items.length; i++) {
			htmlItem = this.render(items[i], this.$output);
			if (htmlItem) {
				html += htmlItem;
			}
		};
		if (html) {
			this.$output.append(html);
		}
	}

	camlBuilder(value) {
		let column, camlObj;
		let caml = '';
		let searchQuery = [];
		if (typeOf(this.columns) === 'String') {
			this.columns = [this.columns];
		}
		for (let j = 0; j < this.columns.length; j++) {
			column = this.columns[j];
			camlObj = {};
			camlObj[column] = {
				Operand: 'Search',
				Value: value,
				Type: this.columnsData[column].TypeAsString
			};
			searchQuery.push(camlObj)
		};
		this.camlSearch = value && this.spx.getCamlQuery({
			OR: searchQuery
		});
		return this.camlSearch
	}

	concatCaml(caml) {
		let sourceCaml = typeOf(this.caml) === 'object' ? this.spx.getCamlQuery(this.caml) : this.spx.parseCamlString(this.caml);
		if (sourceCaml) {
			if (caml) {
				return '<And>' + sourceCaml + caml + '</And>'
			} else {
				return sourceCaml
			}
		} else {
			if (caml) {
				return caml
			} else {
				return ''
			}
		}
	}

	setCaml(caml) {
		let it = this;
		this._pager && this._pager.destroy();
		it.$output.html('');
		it.caml = caml;
		this.auraForm.inputElements[0].getValue();
	}

	getSearchCaml() {
		return this.camlSearch;
	}

	setText(queryText) {
		if (typeOf(queryText) === 'string') {
			this.auraForm.inputElements[0].fill({
				value: queryText.decode()
			})
		}
		return false
	}

	destroy() {
		this._pager && this._pager.destroy();
	}
}