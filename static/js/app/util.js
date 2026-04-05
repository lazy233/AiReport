/**
 * 全局工具（避免与后续打包方案冲突，暂挂到 window.PptApp）
 */
window.PptApp = window.PptApp || {};

PptApp.sleep = function (ms) {
  return new Promise(function (resolve) {
    setTimeout(resolve, ms);
  });
};
