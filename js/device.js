document.addEventListener("DOMContentLoaded", function () {
  const el = document.getElementById("device-info");

  if (!el) return;

  const ua = navigator.userAgent;

  let device = "Unknown device";

  if (/android/i.test(ua)) {
    device = "Android device";
  } else if (/iPhone/i.test(ua)) {
    device = "iPhone";
  } else if (/iPad/i.test(ua)) {
    device = "iPad";
  } else if (/Mac/i.test(ua)) {
    device = "Mac";
  } else if (/Windows/i.test(ua)) {
    device = "Windows PC";
  }

  el.innerText = device;
});