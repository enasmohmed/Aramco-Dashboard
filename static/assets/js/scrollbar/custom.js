(function () {
  var myElement = document.getElementById("simple-bar");
  // بعض الصفحات لا تحتوي على #simple-bar (مثل صفحات بLayout مختلف) — تجنب تكسير باقي السكربتات
  if (!myElement || typeof SimpleBar === "undefined") return;
  new SimpleBar(myElement, { autoHide: true });
})();
