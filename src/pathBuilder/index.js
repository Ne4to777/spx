export default class pathBuilder {
  constructor() {
    this.pathMap = new Map
  }
  get(key) {
    return this.pathMap.get(key);
  }
  set(key, value) {
    let sanitizedPath = this.sanitize(value);
    sanitizedPath && this.pathMap.set(key, sanitizedPath);
  }
  build() {
    return [...this.pathMap.values()].join('/');
  }
  sanitize(path) {
    if (/^\//.test(path)) path = path.substring(1);
    if (/\/$/.test(path)) path = path.substring(0, path.length - 2);
    return path;
  }
}