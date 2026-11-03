const sections = [...document.querySelectorAll('.research-section')];
const navLinks = [...document.querySelectorAll('.sidebar nav a')];
let scheduled=false;
function updateActive(){
 scheduled=false;
 const current=[...sections].reverse().find(s=>!s.hidden&&s.getBoundingClientRect().top<=100)||sections.find(s=>!s.hidden);
 navLinks.forEach(a=>{const active=a.hash==='#'+current?.id;a.classList.toggle('active',active);if(active)a.setAttribute('aria-current','location');else a.removeAttribute('aria-current');});
}
addEventListener('scroll',()=>{if(!scheduled){requestAnimationFrame(updateActive);scheduled=true;}},{passive:true});
updateActive();

