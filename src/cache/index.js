const cache = {};
export const get = paths => {
  const filteredPaths = paths.filter(el => !!el.length);
  const getR = (base, i) => {
    if (base) {
      const basePath = base[filteredPaths[i]];
      return i < filteredPaths.length - 1 ? getR(basePath, ++i) : basePath;
    } else {
      return null;
    }
  }
  return getR(cache, 0);
}
export const set = (paths, spObjects) => {
  const filteredPaths = paths.filter(el => !!el.length);
  const setR = (base, i) => {
    const path = filteredPaths[i];
    if (i < filteredPaths.length - 1) {
      if (!base[path]) base[path] = {}
      setR(base[path], ++i);
    } else {
      base[path] = spObjects;
    }
  }
  setR(cache, 0);
}
export const unset = paths => {
  const filteredPaths = paths.filter(el => !!el.length);
  const deleteR = (base, i) => {
    if (base) {
      const path = filteredPaths[i];
      if (i < filteredPaths.length - 1) {
        const basePath = base[path];
        if (basePath) deleteR(basePath, ++i);
      } else {
        delete base[path];
      }
    }
  }
  deleteR(cache, 0);
}
export const show = () => console.log(cache);

export const getCache = () => cache;