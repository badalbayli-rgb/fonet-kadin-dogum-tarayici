(() => {
  const version = '1.4.0';
  const base = 'https://badalbayli-rgb.github.io/fonet-kadin-dogum-tarayici/hasta-excel-tarayici/';
  const load = (url, done) => {
    const script = document.createElement('script');
    script.src = url;
    script.onload = done || null;
    script.onerror = () => alert('FONET tarayıcı dosyası yüklenemedi. İnternet bağlantısını kontrol edin.');
    document.documentElement.appendChild(script);
  };
  const start = () => load(`${base}scanner.js?v=${version}-${Date.now()}`);
  if (window.XLSX) start();
  else load('https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js', start);
})();
