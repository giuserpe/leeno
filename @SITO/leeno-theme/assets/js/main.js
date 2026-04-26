/**
 * LeenO Digital Masonry — Main JavaScript
 * WebGL Concrete Shader + CAD Grid + GSAP Animations
 */

(function () {
    'use strict';

    const themeUrl = window.leenoData ? window.leenoData.themeUrl : '';

    /* ============================================================
       1. WEBGL CONCRETE DISPLACEMENT SHADER (Hero)
       ============================================================ */
    function initHeroShader() {
        const canvas = document.getElementById('heroCanvas');
        if (!canvas) return;

        const gl = canvas.getContext('webgl', { antialias: false, alpha: false });
        if (!gl) return;

        const VERT = `
            attribute vec2 a_position;
            varying vec2 v_uv;
            void main() {
                v_uv = a_position * 0.5 + 0.5;
                gl_Position = vec4(a_position, 0.0, 1.0);
            }
        `;

        const FRAG = `
            precision highp float;
            uniform float u_time;
            uniform vec2 u_mouse;
            uniform vec2 u_resolution;
            uniform sampler2D u_texture;
            varying vec2 v_uv;

            #define PI 3.14159265359

            float hash(vec2 p) {
                vec3 p3 = fract(vec3(p.xyx) * 0.1031);
                p3 += dot(p3, p3.yzx + 33.33);
                return fract((p3.x + p3.y) * p3.z);
            }

            float noise(vec2 p) {
                vec2 i = floor(p);
                vec2 f = fract(p);
                f = f * f * (3.0 - 2.0 * f);
                return mix(
                    mix(hash(i), hash(i + vec2(1.0, 0.0)), f.x),
                    mix(hash(i + vec2(0.0, 1.0)), hash(i + vec2(1.0, 1.0)), f.x),
                    f.y
                );
            }

            float fbm(vec2 p) {
                float value = 0.0;
                float amplitude = 0.5;
                float frequency = 3.0;
                for (int i = 0; i < 5; i++) {
                    value += amplitude * noise(p * frequency);
                    amplitude *= 0.5;
                    frequency *= 2.0;
                    p = mat2(0.8, 0.6, -0.6, 0.8) * p;
                }
                return value;
            }

            float concreteTexture(vec2 uv) {
                float baseTex = texture2D(u_texture, uv * 4.0).r;
                return baseTex * 0.7 + fbm(uv * 3.0) * 0.3;
            }

            float displacement(vec2 uv, float t) {
                float disp = concreteTexture(uv) * 0.08;
                disp += sin(uv.y * 8.0 + t * 0.5) * 0.008;
                disp += noise(uv * 5.0 + t * 0.2) * 0.01;
                return disp;
            }

            void main() {
                vec2 uv = v_uv;
                float t = u_time * 0.3;
                float mouseDist = length(uv - u_mouse);
                float mouseInfluence = smoothstep(0.5, 0.0, mouseDist) * 0.15;
                vec2 displacedUV = uv + displacement(uv, t) + mouseInfluence;
                float texValue = concreteTexture(displacedUV);
                float edge = fbm(uv * 2.0) * 0.4 + 0.3;
                float finalColor = texValue * edge;
                finalColor *= (1.0 - mouseInfluence * 2.0);
                finalColor = pow(finalColor, 0.9);
                vec3 color = mix(vec3(0.05, 0.05, 0.05), vec3(0.07, 0.1, 0.14), smoothstep(0.1, 0.6, finalColor));
                float cyanGlow = exp(-mouseDist * mouseDist * 20.0) * 0.4;
                color += vec3(0.0, 0.9, 1.0) * cyanGlow;
                gl_FragColor = vec4(color, 1.0);
            }
        `;

        function compile(type, src) {
            const s = gl.createShader(type);
            gl.shaderSource(s, src);
            gl.compileShader(s);
            return s;
        }

        const vs = compile(gl.VERTEX_SHADER, VERT);
        const fs = compile(gl.FRAGMENT_SHADER, FRAG);
        const prog = gl.createProgram();
        gl.attachShader(prog, vs);
        gl.attachShader(prog, fs);
        gl.linkProgram(prog);
        gl.useProgram(prog);

        // Quad
        const buf = gl.createBuffer();
        gl.bindBuffer(gl.ARRAY_BUFFER, buf);
        gl.bufferData(gl.ARRAY_BUFFER, new Float32Array([
            -1,-1, 1,-1, -1,1,
            -1,1, 1,-1, 1,1
        ]), gl.STATIC_DRAW);
        const pos = gl.getAttribLocation(prog, 'a_position');
        gl.enableVertexAttribArray(pos);
        gl.vertexAttribPointer(pos, 2, gl.FLOAT, false, 0, 0);

        const uTime = gl.getUniformLocation(prog, 'u_time');
        const uMouse = gl.getUniformLocation(prog, 'u_mouse');
        const uRes = gl.getUniformLocation(prog, 'u_resolution');
        const uTex = gl.getUniformLocation(prog, 'u_texture');

        // Texture
        const tex = gl.createTexture();
        gl.bindTexture(gl.TEXTURE_2D, tex);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_S, gl.REPEAT);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_WRAP_T, gl.REPEAT);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MIN_FILTER, gl.LINEAR);
        gl.texParameteri(gl.TEXTURE_2D, gl.TEXTURE_MAG_FILTER, gl.LINEAR);

        // Placeholder pixel
        gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, 1, 1, 0, gl.RGBA, gl.UNSIGNED_BYTE, new Uint8Array([128,128,128,255]));

        // Load texture
        const img = new Image();
        img.crossOrigin = 'anonymous';
        img.onload = function () {
            gl.bindTexture(gl.TEXTURE_2D, tex);
            gl.texImage2D(gl.TEXTURE_2D, 0, gl.RGBA, gl.RGBA, gl.UNSIGNED_BYTE, img);
        };
        img.src = themeUrl + '/assets/images/concrete-texture.jpg';

        const mouse = { x: 0.5, y: 0.5 };
        const targetMouse = { x: 0.5, y: 0.5 };

        function onMove(e) {
            targetMouse.x = e.clientX / window.innerWidth;
            targetMouse.y = 1.0 - e.clientY / window.innerHeight;
        }
        window.addEventListener('mousemove', onMove);

        function resize() {
            const dpr = Math.min(window.devicePixelRatio, 1.5);
            canvas.width = canvas.offsetWidth * dpr;
            canvas.height = canvas.offsetHeight * dpr;
            gl.viewport(0, 0, canvas.width, canvas.height);
        }
        window.addEventListener('resize', resize);
        resize();

        const start = performance.now();
        let raf;

        function loop() {
            raf = requestAnimationFrame(loop);
            mouse.x += (targetMouse.x - mouse.x) * 0.05;
            mouse.y += (targetMouse.y - mouse.y) * 0.05;
            const elapsed = (performance.now() - start) / 1000;
            gl.uniform1f(uTime, elapsed);
            gl.uniform2f(uMouse, mouse.x, mouse.y);
            gl.uniform2f(uRes, canvas.width, canvas.height);
            gl.uniform1i(uTex, 0);
            gl.drawArrays(gl.TRIANGLES, 0, 6);
        }
        loop();

        // Cleanup on page hide
        document.addEventListener('visibilitychange', function () {
            if (document.hidden) cancelAnimationFrame(raf);
            else loop();
        });
    }

    /* ============================================================
       2. CAD GRID (Canvas 2D)
       ============================================================ */
    function initCADGrid(canvasId, opts) {
        const canvas = document.getElementById(canvasId);
        if (!canvas) return;

        const ctx = canvas.getContext('2d');
        const curved = opts && opts.curved;

        const CELL = 40;
        const STEPS = 20;
        const MAX_LINES = 100;
        const PROB = 0.005;
        const FADE = 0.015;

        let points = [];
        let lines = [];
        let raf;

        function buildGrid() {
            points = [];
            lines = [];
            const w = canvas.width;
            const h = canvas.height;
            const dpr = Math.min(window.devicePixelRatio, 1.5);
            const cw = w / dpr;
            const ch = h / dpr;
            for (let x = 0; x <= cw; x += CELL) {
                for (let y = 0; y <= ch; y += CELL) {
                    points.push({ x, y, connections: [] });
                }
            }
        }

        function resize() {
            const dpr = Math.min(window.devicePixelRatio, 1.5);
            const rect = canvas.parentElement.getBoundingClientRect();
            canvas.width = rect.width * dpr;
            canvas.height = rect.height * dpr;
            ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
            buildGrid();
        }

        function addLine() {
            if (lines.length >= MAX_LINES) return;
            const start = points[Math.floor(Math.random() * points.length)];
            const neighbors = [];
            for (let i = 0; i < points.length; i++) {
                const p = points[i];
                if (Math.hypot(p.x - start.x, p.y - start.y) === CELL) {
                    neighbors.push(p);
                }
            }
            if (neighbors.length === 0) return;
            const end = neighbors[Math.floor(Math.random() * neighbors.length)];
            if (start.connections.some(function (c) { return c.id === end.x + '-' + end.y; })) return;

            const cp = [{ x: start.x, y: start.y }];
            for (let i = 1; i < STEPS; i++) {
                const t = i / STEPS;
                cp.push({ x: start.x + (end.x - start.x) * t, y: start.y + (end.y - start.y) * t });
            }
            cp.push({ x: end.x, y: end.y });

            if (curved) {
                for (let i = 1; i < cp.length - 1; i++) {
                    const n = (Math.random() - 0.5) * (CELL * 0.8);
                    cp[i].x += n;
                    cp[i].y += n;
                }
            }

            const straight = Math.random() > 0.7;
            const line = {
                points: straight ? [{ x: start.x, y: start.y }, { x: end.x, y: end.y }] : cp,
                progress: 0,
                speed: 1.5 + Math.random() * 1.5,
                node1: start,
                node2: end,
            };
            lines.push(line);
            start.connections.push({ id: end.x + '-' + end.y });
            end.connections.push({ id: start.x + '-' + start.y });
        }

        function loop() {
            raf = requestAnimationFrame(loop);
            const dpr = Math.min(window.devicePixelRatio, 1.5);
            const cw = canvas.width / dpr;
            const ch = canvas.height / dpr;

            ctx.fillStyle = 'rgba(18, 26, 35, ' + FADE + ')';
            ctx.fillRect(0, 0, cw, ch);

            // Nodes
            for (let i = 0; i < points.length; i++) {
                const p = points[i];
                if (Math.random() < 0.02) {
                    ctx.strokeStyle = 'rgba(0, 229, 255, 0.3)';
                    ctx.lineWidth = 1;
                    ctx.beginPath(); ctx.moveTo(p.x - 2, p.y); ctx.lineTo(p.x + 2, p.y); ctx.stroke();
                    ctx.beginPath(); ctx.moveTo(p.x, p.y - 2); ctx.lineTo(p.x, p.y + 2); ctx.stroke();
                }
                for (let j = 0; j < p.connections.length; j++) {
                    const conn = p.connections[j];
                    ctx.strokeStyle = 'rgba(255, 255, 255, 0.04)';
                    ctx.beginPath(); ctx.moveTo(p.x, p.y);
                    const match = lines.find(function (l) { return l.node2.x + '-' + l.node2.y === conn.id; });
                    if (match) ctx.lineTo(match.node2.x, match.node2.y);
                    else ctx.lineTo(p.x, p.y);
                    ctx.stroke();
                }
            }

            if (Math.random() < PROB) addLine();

            for (let i = lines.length - 1; i >= 0; i--) {
                const line = lines[i];
                const ti = Math.floor(line.progress);
                if (ti >= line.points.length - 1) { lines.splice(i, 1); continue; }
                const t = line.progress - ti;
                const ease = t < 0.5 ? 4 * t * t * t : 1 - Math.pow(-2 * t + 2, 3) / 2;
                const p0 = line.points[ti];
                const p1 = line.points[ti + 1];
                const x = p0.x + (p1.x - p0.x) * ease;
                const y = p0.y + (p1.y - p0.y) * ease;
                ctx.strokeStyle = 'rgba(0, 229, 255, 0.7)';
                ctx.lineWidth = 1;
                ctx.beginPath(); ctx.moveTo(p0.x, p0.y); ctx.lineTo(x, y); ctx.stroke();
                line.progress += line.speed / STEPS;
            }
        }

        resize();
        window.addEventListener('resize', resize);
        loop();
    }

    /* ============================================================
       3. GSAP ANIMATIONS
       ============================================================ */
    function initAnimations() {
        if (typeof gsap === 'undefined' || typeof ScrollTrigger === 'undefined') return;
        gsap.registerPlugin(ScrollTrigger);

        // Header scroll
        const header = document.getElementById('siteHeader');
        if (header) {
            window.addEventListener('scroll', function () {
                header.classList.toggle('scrolled', window.scrollY > 80);
            }, { passive: true });
        }

        // Hero title entrance
        const heroTitle = document.getElementById('heroTitle');
        if (heroTitle) {
            const words = heroTitle.querySelectorAll('.word-inner');
            gsap.from(words, {
                y: 120,
                opacity: 0,
                duration: 1.2,
                stagger: 0.08,
                ease: 'back.out(1.2)',
                delay: 0.3,
            });
        }

        const heroSub = document.getElementById('heroSubtitle');
        if (heroSub) {
            gsap.from(heroSub, {
                y: 30,
                opacity: 0,
                duration: 0.8,
                ease: 'power3.out',
                delay: 1.2,
            });
        }

        // Transition text
        const transText = document.getElementById('transitionText');
        if (transText) {
            gsap.from(transText, {
                opacity: 0,
                y: 60,
                duration: 1,
                ease: 'power3.out',
                scrollTrigger: {
                    trigger: transText.parentElement,
                    start: 'top 80%',
                    end: 'top 20%',
                    scrub: 1,
                },
            });
        }

        // Features title reveal
        const featTitle = document.getElementById('featuresTitle');
        if (featTitle) {
            const inners = featTitle.querySelectorAll('.title-word-inner');
            const outers = featTitle.querySelectorAll('.title-word-outer');
            const tl = gsap.timeline({
                scrollTrigger: {
                    trigger: featTitle,
                    start: 'top 70%',
                    toggleActions: 'play none none none',
                },
            });
            tl.from(inners, { y: 80, duration: 1, stagger: 0.05, ease: 'back.out(1.2)' });
            tl.to(outers, { paddingRight: 15, duration: 1, ease: 'power3.out' }, '-=0.6');
        }

        // Feature cards
        const featCards = document.querySelectorAll('.feature-card');
        if (featCards.length) {
            gsap.from(featCards, {
                y: 60,
                opacity: 0,
                duration: 0.8,
                stagger: 0.15,
                ease: 'power3.out',
                scrollTrigger: {
                    trigger: featCards[0].parentElement,
                    start: 'top 75%',
                    toggleActions: 'play none none none',
                },
            });
        }

        // Screenshot
        const screenshot = document.querySelector('.screenshot-wrap');
        if (screenshot) {
            gsap.from(screenshot, {
                y: 40,
                opacity: 0,
                duration: 0.8,
                ease: 'power3.out',
                scrollTrigger: {
                    trigger: screenshot,
                    start: 'top 80%',
                    toggleActions: 'play none none none',
                },
            });
        }

        // Metrics heading
        const metHead = document.querySelector('.metrics-heading');
        if (metHead) {
            gsap.from(metHead, {
                y: 40,
                opacity: 0,
                duration: 0.8,
                ease: 'power3.out',
                scrollTrigger: {
                    trigger: metHead,
                    start: 'top 70%',
                    toggleActions: 'play none none none',
                },
            });
        }

        // Blog section title
        const blogTitle = document.querySelector('.blog-section .section-title');
        if (blogTitle) {
            const inners = blogTitle.querySelectorAll('.title-word-inner');
            const outers = blogTitle.querySelectorAll('.title-word-outer');
            const tl = gsap.timeline({
                scrollTrigger: {
                    trigger: blogTitle,
                    start: 'top 70%',
                    toggleActions: 'play none none none',
                },
            });
            tl.from(inners, { y: 80, duration: 1, stagger: 0.05, ease: 'back.out(1.2)' });
            tl.to(outers, { paddingRight: 15, duration: 1, ease: 'power3.out' }, '-=0.6');
        }

        // Blog cards
        const postCards = document.querySelectorAll('.post-card');
        if (postCards.length) {
            gsap.from(postCards, {
                y: 60,
                opacity: 0,
                duration: 0.8,
                stagger: 0.1,
                ease: 'power3.out',
                scrollTrigger: {
                    trigger: postCards[0].parentElement,
                    start: 'top 75%',
                    toggleActions: 'play none none none',
                },
            });
        }

        // Footer CTA cards
        const footerCards = document.querySelectorAll('.footer-cta-card');
        if (footerCards.length) {
            gsap.from(footerCards, {
                y: 40,
                opacity: 0,
                duration: 0.8,
                stagger: 0.15,
                ease: 'power3.out',
                scrollTrigger: {
                    trigger: footerCards[0].parentElement,
                    start: 'top 80%',
                    toggleActions: 'play none none none',
                },
            });
        }
    }

    /* ============================================================
       4. FLIP-STAT COUNTERS
       ============================================================ */
    function initFlipStats() {
        const stats = document.querySelectorAll('.flip-stat');
        if (!stats.length) return;

        stats.forEach(function (container) {
            const finalValue = container.textContent.trim();
            container.innerHTML = '';

            const digitObjects = [];
            finalValue.split('').forEach(function (ch) {
                const span = document.createElement('span');
                span.style.display = 'inline-block';
                span.style.minWidth = '0.55em';
                span.textContent = ch === '.' ? '.' : (ch === ' ' ? '\u00a0' : '0');
                if (ch !== ' ' && ch !== '.') {
                    span.dataset.final = ch;
                }
                container.appendChild(span);
                digitObjects.push({ el: span, final: ch });
            });

            function setFinal(digit) {
                digit.textContent = digit.dataset.final || '0';
            }

            function scramble() {
                digitObjects.forEach(function (obj) {
                    if (obj.final === ' ' || obj.final === '.') return;
                    setTimeout(function () { setFinal(obj.el); }, 800);
                    const interval = setInterval(function () {
                        obj.el.textContent = String(Math.floor(Math.random() * 10));
                    }, 20);
                    setTimeout(function () { clearInterval(interval); }, 800);
                });
            }

            const observer = new IntersectionObserver(function (entries) {
                entries.forEach(function (entry) {
                    if (entry.isIntersecting) {
                        scramble();
                        observer.disconnect();
                    }
                });
            }, { threshold: 0.5 });

            observer.observe(container);
        });
    }

    /* ============================================================
       5. MOBILE MENU + SUBMENU ACCORDION
       ============================================================ */
    function initMobileMenu() {
        const toggle = document.getElementById('menuToggle');
        const nav = document.getElementById('mainNav');
        if (!toggle || !nav) return;

        toggle.addEventListener('click', function () {
            nav.classList.toggle('active');
            toggle.classList.toggle('open');
        });

        // Mobile submenu accordion
        const hasChildren = nav.querySelectorAll('.menu-item-has-children');
        hasChildren.forEach(function (item) {
            const link = item.querySelector(':scope > a');
            if (!link) return;

            link.addEventListener('click', function (e) {
                // Only on mobile when nav is active
                if (!nav.classList.contains('active')) return;
                e.preventDefault();
                item.classList.toggle('submenu-open');
            });
        });

        // Close mobile menu on window resize to desktop
        window.addEventListener('resize', function () {
            if (window.innerWidth > 768 && nav.classList.contains('active')) {
                nav.classList.remove('active');
                toggle.classList.remove('open');
            }
        });
    }

    /* ============================================================
       INIT ALL
       ============================================================ */
    document.addEventListener('DOMContentLoaded', function () {
        initHeroShader();
        initCADGrid('cadCanvas', { curved: false });
        initCADGrid('cadCanvasCurved', { curved: true });
        initAnimations();
        initFlipStats();
        initMobileMenu();
    });

})();

/* ============================================================
   IMPROVEMENTS v1.1
   ============================================================ */

/* ——— Contatore metriche con IntersectionObserver ——— */
function initMetricCounters() {
    // Seleziona TUTTE le metric-item, con o senza data-target
    const allItems  = document.querySelectorAll('.metric-item');
    const countItems = document.querySelectorAll('.metric-item[data-target]');
    if (!allItems.length) return;

    function easeOutExpo(t) {
        return t === 1 ? 1 : 1 - Math.pow(2, -10 * t);
    }

    function animateCounter(el, target, duration) {
        const start = performance.now();
        const stat = el.querySelector('.flip-stat');
        if (!stat) return;

        function update(now) {
            const elapsed = now - start;
            const progress = Math.min(elapsed / duration, 1);
            const eased = easeOutExpo(progress);
            const current = Math.round(eased * target);
            stat.textContent = current.toLocaleString('it-IT');
            if (progress < 1) requestAnimationFrame(update);
        }
        requestAnimationFrame(update);
    }

    const observer = new IntersectionObserver((entries) => {
        entries.forEach(entry => {
            if (entry.isIntersecting && !entry.target.dataset.counted) {
                entry.target.dataset.counted = '1';
                // Sempre: rivela con animazione CSS
                entry.target.classList.add('is-visible');
                // Solo se ha un target numerico: anima il contatore
                const target = parseInt(entry.target.dataset.target, 10);
                if (!isNaN(target)) {
                    const duration = target > 100 ? 1800 : 1200;
                    animateCounter(entry.target, target, duration);
                }
            }
        });
    }, { threshold: 0.3 });

    // Osserva tutte le card, non solo quelle con data-target
    allItems.forEach(item => observer.observe(item));
}

/* ——— Filtro live per WP Filebase ——— */
function initPrezzariFilter() {
    const input = document.getElementById('prezzari-filter');
    if (!input) return;

    const rows   = document.querySelectorAll('.prezzari-row');
    const blocks = document.querySelectorAll('.prezzari-block');

    input.addEventListener('input', function () {
        const q = this.value.toLowerCase().trim();

        rows.forEach(row => {
            const text = row.textContent.toLowerCase();
            row.classList.toggle('is-hidden', q !== '' && !text.includes(q));
        });

        // Nascondi blocco regione/sottocategoria se non ha righe visibili
        // Scorri dal basso verso l'alto per gestire la gerarchia
        Array.from(blocks).reverse().forEach(block => {
            if (!q) {
                block.classList.remove('is-hidden');
                return;
            }
            const visibleRows = block.querySelectorAll('.prezzari-row:not(.is-hidden)');
            const visibleSubcats = block.querySelectorAll('.prezzari-subcat:not(.is-hidden)');
            block.classList.toggle('is-hidden', visibleRows.length === 0 && visibleSubcats.length === 0);
        });
    });
}

/* ——— Scroll smooth per "Scopri le funzioni" ——— */
function initHeroScrollCta() {
    const cta = document.getElementById('heroScrollCta');
    if (!cta) return;
    cta.addEventListener('click', function (e) {
        const target = document.querySelector(this.getAttribute('href'));
        if (target) {
            e.preventDefault();
            target.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }
    });
}

/* ——— Aria expanded sul menu mobile ——— */
function initMenuToggleAria() {
    const toggle = document.getElementById('menuToggle');
    const nav = document.getElementById('mainNav');
    if (!toggle || !nav) return;
    const origClick = toggle.onclick;
    toggle.addEventListener('click', function () {
        const expanded = toggle.getAttribute('aria-expanded') === 'true';
        toggle.setAttribute('aria-expanded', String(!expanded));
        toggle.setAttribute('aria-label', !expanded ? 'Chiudi menu' : 'Apri menu');
    });
}

/* ——— Header search toggle ——— */
function initHeaderSearch() {
    const toggle = document.getElementById('headerSearchToggle');
    const box    = document.getElementById('headerSearchBox');
    const input  = document.getElementById('headerSearchInput');
    if (!toggle || !box) return;

    toggle.addEventListener('click', function (e) {
        e.stopPropagation();
        const isOpen = !box.hidden;
        box.hidden = isOpen;
        toggle.setAttribute('aria-expanded', String(!isOpen));
        if (!isOpen && input) input.focus();
    });

    document.addEventListener('click', function (e) {
        if (!box.hidden && !box.contains(e.target) && e.target !== toggle) {
            box.hidden = true;
            toggle.setAttribute('aria-expanded', 'false');
        }
    });

    document.addEventListener('keydown', function (e) {
        if (e.key === 'Escape' && !box.hidden) {
            box.hidden = true;
            toggle.setAttribute('aria-expanded', 'false');
            toggle.focus();
        }
    });
}

/* ——— Init — unico DOMContentLoaded ——— */
document.addEventListener('DOMContentLoaded', function () {
    initMetricCounters();
    initPrezzariFilter();
    initHeroScrollCta();
    initMenuToggleAria();
    initHeaderSearch();
});
