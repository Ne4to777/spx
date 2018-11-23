export default class Pager {
	constructor() {
		const it = this;
		const viewFields = ['ID'];
		this.$window = $(window);
		this.spx = new SPX;
		this.index = 0;
		this.id = CryptoJS.MD5(new Date().getTime().toString()).toString();
		this.listData = typeOf(listData) === 'string' ? this.spx.getListData(listData) : listData;
		this.opts = opts || {};
		this.opts.isRaw = true;
		this.caml = this.opts.caml;
		this.position = null;
		this.pages = [];
		this.pageSize = this.opts.pageSize || 20;
		this.pageStart = this.opts.pageStart || 1;
		this.recursive = this.opts.recursive || false;
		this.orderByDate = this.opts.orderByDate;
		this.preload = this.opts.preload;
		this.orderBy = this.opts.orderBy || 'ID';
		this.isScrollable = this.opts.isScrollable;
		this.cb = this.opts.cb;
		this.ascending = this.opts.ascending === void 0 ? true : this.opts.ascending;
		this.pending = false;
		this.cache = {
			isNext: {},
			isPrevious: {}
		}
		this.isBusy = {};
		this.pagesPositions = {
			next: null,
			previous: null,
			isLast: null
		};
		if (this.opts.mapBy) {
			viewFields.push(this.opts.mapBy);
		}
		if (this.orderBy) {
			viewFields.push(this.orderBy);
		}
		if (this.opts.viewFields && this.opts.viewFields.length) {
			viewFields.push(this.orderBy);
		}
		if (this.opts.viewFields && this.opts.viewFields.length) {
			this.opts.viewFields = viewFields.concat(this.opts.viewFields);
		}
		if (this.opts.isAsync) {
			this.isAsync = this.opts.isAsync;
		}
		if (this.isScrollable) {
			fillPageR();
			this.$window.off('scroll.pager_' + this.id).on('scroll.pager_' + this.id, function() {
				if (!it.pending && (it.$window.scrollTop() + it.$window.height() + 1 >= $(document).height())) {
					it.moveNext();
				}
			});
		}
	}

	fillPageR() {
		const it = this;
		it.moveNext(function(data) {
			if (it.isAsync) {
				const deferredSend = $.Deferred();
				it.cb(data, deferredSend);
				$.when(deferredSend).then(function() {
					if (data && data.length === it.pageSize && (it.$window.height() + 1 >= $(document).height())) {
						it.fillPageR();
					}
				})
			} else {
				it.cb(data);
				if (data && data.length === it.pageSize && (it.$window.height() + 1 >= $(document).height())) {
					it.fillPageR();
				}
			}
		})
	}

	reset() {
		this.position = null;
		this.pagesPositions = {
			next: null,
			previous: null,
			isLast: null
		};
		this.opts.caml = this.caml;
	}
	moveNext(cb) {
		const nextBase = this.cache.isNext;
		const busyBase = this.isBusy;
		const it = this;
		const key = this.id;
		const cachePath = [this.listData.path, this.listData.title, key, 'next'];
		if (!this.spx.checkNodeExists.call(this, nextBase, cachePath)) {
			this.spx.createNodeTree.call(this, nextBase, cachePath);
			nextBase[this.listData.path][this.listData.title][key].next = [];
		}
		if (this.spx.checkNodeExists.call(this, busyBase, cachePath)) {
			nextBase[this.listData.path][this.listData.title][key].next.push({
				key: key,
				cb: cb
			})
		} else {
			this.spx.createNodeTree.call(this, busyBase, cachePath);
			moveR(this.listData, key, cb);
		}

		function moveR(listData, id, cb) {
			const cacheBusy = busyBase[listData.path][listData.title][id];
			cacheBusy.next = true;
			it.move('next', function(data) {
				let next;
				let nexts = nextBase[listData.path][listData.title][id].next;
				cacheBusy.next = false;
				if (nexts.length) {
					next = nexts.shift();
					if (it.isAsync) {
						const deferredSend = $.Deferred();
						cb ? cb(data, deferredSend) : it.cb(data, deferredSend);
						$.when(deferredSend).then(function() {
							moveR(listData, next.key, next.cb);
						})
					} else {
						cb ? cb(data) : it.cb(data);
						moveR(listData, next.key, next.cb);
					}
				} else {
					if (it.isAsync) {
						const deferredSend = $.Deferred();
						cb ? cb(data, deferredSend, ++it.index) : it.cb(data, deferredSend, ++it.index);
						$.when(deferredSend).then(function() {})
					} else {
						cb ? cb(data, ++it.index) : it.cb(data, ++it.index);
					}
				}
			});
		}
	}

	movePrevious(cb) {
		const previousBase = this.cache.isPrevious;
		const busyBase = this.isBusy;
		const it = this;
		const key = this.id;
		const cachePath = [this.listData.path, this.listData.title, key, 'previous'];
		if (!this.spx.checkNodeExists.call(this, previousBase, cachePath)) {
			this.spx.createNodeTree.call(this, previousBase, cachePath);
			previousBase[this.listData.path][this.listData.title][key].previous = [];
		}
		if (this.spx.checkNodeExists.call(this, busyBase, cachePath)) {
			previousBase[this.listData.path][this.listData.title][key].previous.push({
				key: key,
				cb: cb
			})
		} else {
			this.spx.createNodeTree.call(this, busyBase, cachePath);
			moveR(this.listData, key, cb);
		}

		function moveR(listData, id, cb) {
			const cacheBusy = busyBase[listData.path][listData.title][id];
			cacheBusy.previous = true;
			it.move('previous', function(data) {
				let previous;
				let previouses = previousBase[listData.path][listData.title][id].previous;
				cacheBusy.previous = false;
				if (previouses.length) {
					previous = previouses.shift();
					if (it.isAsync) {
						const deferredSend = $.Deferred();
						cb ? cb(data, deferredSend) : it.cb(data, deferredSend);
						$.when(deferredSend).then(function() {
							moveR(listData, previous.key, previous.cb);
						})
					} else {
						cb ? cb(data) : it.cb(data);
						moveR(listData, previous.key, previous.cb);
					}
				} else {
					if (it.isAsync) {
						const deferredSend = $.Deferred();
						cb ? cb(data, deferredSend, --it.index) : it.cb(data, deferredSend, --it.index);
						$.when(deferredSend).then(function() {})
					} else {
						cb ? cb(data, --it.index) : it.cb(data, --it.index);
					}
				}
			});
		}
	}

	move(direction, cb) {
		this.pending = true;
		// console.log(this);
		!direction && (direction = 'next');
		let caml;
		const it = this;
		const pagesPositions = this.pagesPositions;
		const camlQuery = new SP.CamlQuery();
		typeOf(this.preload) === 'function' && this.preload();
		if (typeOf(cb) !== 'function') {
			console.log('cb is not a function');
			return
		};
		if (pagesPositions[direction]) {
			this.position = new SP.ListItemCollectionPosition();
			this.position.set_pagingInfo(pagesPositions[direction]);
			camlQuery.set_listItemCollectionPosition(this.position);
		}
		if (direction === 'next' && pagesPositions.isLast) {
			cb(null);
			return;
		}
		if (this.caml) {
			if (typeOf(this.caml) === 'object') {
				caml = this.spx.getCamlQuery(this.caml);
			} else if (!/\<FieldRef/i.test(this.caml)) {
				caml = this.spx.parseCamlString(this.caml);
			} else {
				caml = this.caml;
			}
		}
		camlQuery.set_viewXml('\
		<View ' + (this.recursive ? 'Scope="' + (this.recursive == 'All' ? 'RecursiveAll' : 'Recursive') + '"' : '') + '>\
			<Query>\
				' + (caml ? ('<Where>' + caml + '</Where>') : '') + '\
				' + (this.orderBy ? '<OrderBy><FieldRef Name="' + this.orderBy + '" Ascending="' + (!!this.ascending).toString().toUpperCase() + '"/></OrderBy>' : '') + '\
			</Query>\
			<RowLimit>' + (this.pageSize * this.pageStart) + '</RowLimit>\
		</View>');
		this.opts.caml = camlQuery;
		it.spx.getListItems(it.listData, function(items) {
			let item0, itemLastId, orderByValue, orderQuery, itemLast, timezoneOffset;
			const itemsLength = items.get_count();
			items.$2_1 && items.$2_1.splice(0, it.pageSize * (it.pageStart - 1));
			if (itemsLength) {
				item0 = items.$2_1[0];
				itemLast = items.$2_1[itemsLength - 1];
				itemLastId = itemLast && itemLast.get_item('ID');
			}
			if (direction === 'previous' && itemsLength < it.pageSize || !itemsLength) {
				it.isScrollable && it.destroy();
				cb(null);
				return;
			} else {
				pagesPositions.isLast = false;
			}
			orderByValue = it.orderBy && itemLast && itemLast.get_item(it.orderBy);
			if (it.orderByDate && orderByValue) {
				timezoneOffset = new Date().getTimezoneOffset() / 60;
				orderByValue.setHours(orderByValue.getHours() + timezoneOffset);
				orderByValue = orderByValue.format('yyyyMMdd HH:mm:ss');
			}
			orderQuery = orderByValue !== void 0 && it.orderBy !== 'ID' ?
				('&p_' + it.orderBy + '=' + encodeURIComponent(typeOf(orderByValue) === 'object' ?
					orderByValue.get_lookupValue() :
					orderByValue)) : '';
			itemLastId ? (pagesPositions.next = 'Paged=TRUE' + orderQuery + '&p_ID=' + itemLastId) : (pagesPositions.isLast = true);
			item0 && (pagesPositions.previous = 'PagedPrev=TRUE&Paged=TRUE' + orderQuery + '&p_ID=' + item0.get_item('ID'));
			cb(it.spx.prepareResponseCollection(items, it.opts));
			it.pending = false;
		}, it.opts);
	}

	destroy() {
		this.$window.off('scroll.pager_' + this.id);
	};
}