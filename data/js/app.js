/* =====================================================================
 * 本地文档浏览器 - 前端逻辑
 * 数据来源优先级:
 *   1. window.__DOC_DATA (data.js 内联, 双击打开 file:// 时可用)
 *   2. fetch("data.json")               (根目录)
 *   3. fetch("assets/data/data.json")   (扫描器原始位置)
 * ===================================================================== */

(function () {
    "use strict";

    // ============ 全局状态 ============
    const state = {
        data: null,                  // 完整数据对象
        currentFolderId: "",          // 当前选中的文件夹 id, "" 表示根
        viewMode: "grid",             // grid | list
        sortBy: "name-asc",
        keyword: "",                  // 搜索关键词
        folderFilter: "",             // 目录过滤
        extFilter: "",                // 后缀筛选, "" 表示全部
        folderMap: new Map(),         // id -> folder node
        fileCountMap: new Map(),      // folder id -> 直接子文件数
    };

    // ============ DOM ============
    const $ = (id) => document.getElementById(id);
    const dom = {
        treeView: $("treeView"),
        sidebarFooter: $("sidebarFooter"),
        breadcrumb: $("breadcrumb"),
        cardGrid: $("cardGrid"),
        emptyState: $("emptyState"),
        folderInfo: $("folderInfo"),
        resultCount: $("resultCount"),
        sortSelect: $("sortSelect"),
        extFilter: $("extFilter"),
        searchInput: $("searchInput"),
        searchClear: $("searchClear"),
        folderFilter: $("folderFilter"),
        viewGrid: $("viewGrid"),
        viewList: $("viewList"),
        expandAll: $("expandAll"),
        collapseAll: $("collapseAll"),
        refreshBtn: $("refreshBtn"),
        rescanBtn: $("rescanBtn"),
        backToTop: $("backToTop"),
        // modal
        previewModal: $("previewModal"),
        modalClose: $("modalClose"),
        previewImg: $("previewImg"),
        previewTitle: $("previewTitle"),
        previewMeta: $("previewMeta"),
        openInBrowserBtn: $("openInBrowserBtn"),
        openFileBtn: $("openFileBtn"),
        copyPathBtn: $("copyPathBtn"),
        openParentBtn: $("openParentBtn"),
        // gallery
        galleryPrev: $("galleryPrev"),
        galleryNext: $("galleryNext"),
        galleryCounter: $("galleryCounter"),
        // lightbox
        imageLightbox: $("imageLightbox"),
        lightboxImg: $("lightboxImg"),
        lightboxClose: $("lightboxClose"),
        // ctx
        cardMenu: $("cardMenu"),
        // loading
        loadingMask: $("loadingMask"),
        loadingText: $("loadingText"),
        firstRunTip: $("firstRunTip"),
        retryLoadBtn: $("retryLoadBtn"),
        // toast
        toast: $("toast"),
    };

    let currentPreviewFile = null;
    let galleryImages = [];      // 当前预览的所有图片
    let galleryIndex = 0;        // 当前图片索引

    // ============ 工具函数 ============
    function showToast(msg, dur = 2000) {
        dom.toast.textContent = msg;
        dom.toast.style.display = "block";
        clearTimeout(showToast._t);
        showToast._t = setTimeout(() => { dom.toast.style.display = "none"; }, dur);
    }

    function escapeHtml(s) {
        if (s == null) return "";
        return String(s)
            .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
    }

    function fileUriFromPath(p) {
        // Windows 路径 D:\foo\bar.pdf -> file:///D:/foo/bar.pdf
        if (!p) return "";
        let s = p.replace(/\\/g, "/");
        if (s.startsWith("//")) {
            return "file:" + s;
        }
        if (/^[a-zA-Z]:/.test(s)) {
            return "file:///" + s;
        }
        if (s.startsWith("/")) {
            return "file://" + s;
        }
        return "file:///" + s;
    }

    function parentDirOf(p) {
        if (!p) return "";
        const idx = Math.max(p.lastIndexOf("\\"), p.lastIndexOf("/"));
        return idx >= 0 ? p.substring(0, idx) : p;
    }

    async function copyText(text) {
        try {
            await navigator.clipboard.writeText(text);
            showToast("已复制: " + text);
        } catch (e) {
            // 降级: 用 textarea
            const ta = document.createElement("textarea");
            ta.value = text;
            ta.style.position = "fixed";
            ta.style.opacity = "0";
            document.body.appendChild(ta);
            ta.select();
            try {
                document.execCommand("copy");
                showToast("已复制: " + text);
            } catch (e2) {
                showToast("复制失败, 请手动选择: " + text, 4000);
            }
            document.body.removeChild(ta);
        }
    }

    // 自然排序
    function naturalCompare(a, b) {
        return String(a).localeCompare(String(b), "zh-CN", { numeric: true, sensitivity: "base" });
    }

    // ============ 数据加载 ============
    async function loadData() {
        // 1) 内联 data.js
        if (window.__DOC_DATA) {
            return window.__DOC_DATA;
        }
        // 2) fetch 根目录 data.json
        const candidates = ["data.json", "assets/data/data.json", "./data.json"];
        let lastErr = null;
        for (const url of candidates) {
            try {
                const resp = await fetch(url, { cache: "no-store" });
                if (resp.ok) {
                    return await resp.json();
                }
                lastErr = new Error("HTTP " + resp.status);
            } catch (e) {
                lastErr = e;
            }
        }
        throw lastErr || new Error("no data");
    }

    // 构建索引
    function buildIndex() {
        const tree = state.data.tree || [];
        const allFiles = state.data.files || [];
        state.folderMap.clear();
        state.fileCountMap.clear();
        state.folderMap.set("", { id: "", name: "全部", children: tree, type: "folder", depth: 0 });
        // 根节点显示总文件数
        state.fileCountMap.set("", allFiles.length);

        function walk(nodes) {
            for (const n of nodes) {
                if (n.type === "folder") {
                    state.folderMap.set(n.id, n);
                    // 递归统计该目录下所有文件数(含子目录)
                    let total = 0;
                    function countFiles(node) {
                        for (const c of (node.children || [])) {
                            if (c.type === "file") total++;
                            else if (c.type === "folder") countFiles(c);
                        }
                    }
                    countFiles(n);
                    state.fileCountMap.set(n.id, total);
                    walk(n.children || []);
                }
            }
        }
        walk(tree);
    }

    // ============ 渲染侧边栏树 ============
    function renderTree() {
        dom.treeView.innerHTML = "";
        const root = state.folderMap.get("");
        // 顶层 "全部" 节点
        const rootRow = createTreeNode({ id: "", name: "📚 全部文件", children: state.data.tree }, 0, true);
        dom.treeView.appendChild(rootRow);

        // 自动展开第一层
        const firstLevel = rootRow.querySelector(".tree-children");
        if (firstLevel) firstLevel.classList.add("open");

        updateSidebarFooter();
    }

    function createTreeNode(node, depth, isRoot) {
        const wrap = document.createElement("div");
        wrap.className = "tree-node";
        wrap.dataset.id = node.id || "";

        const row = document.createElement("div");
        row.className = "tree-row";
        row.dataset.folderId = node.id || "";
        if (state.currentFolderId === (node.id || "")) row.classList.add("active");

        const children = node.children || [];
        const hasChildren = children.some(c => c.type === "folder");
        const fileCount = state.fileCountMap.get(node.id) || 0;

        // 展开/折叠箭头
        const toggle = document.createElement("span");
        toggle.className = "tree-toggle";
        if (hasChildren) {
            toggle.textContent = "▶";
        } else {
            toggle.classList.add("empty");
        }

        // 图标
        const icon = document.createElement("span");
        icon.className = "tree-icon";
        icon.textContent = isRoot ? "🏠" : (depth === 0 ? "📁" : (hasChildren ? "📂" : "📁"));

        // 名称
        const label = document.createElement("span");
        label.className = "tree-label";
        label.textContent = isRoot ? "全部" : node.name;
        label.title = node.name || "";

        // 计数
        const count = document.createElement("span");
        count.className = "tree-count";
        count.textContent = fileCount > 0 ? fileCount : "";

        row.appendChild(toggle);
        row.appendChild(icon);
        row.appendChild(label);
        row.appendChild(count);

        // 点击行: 切换当前文件夹
        row.addEventListener("click", (e) => {
            if (e.target === toggle) return;
            selectFolder(node.id || "");
            // 顺便展开
            const childrenEl = wrap.querySelector(":scope > .tree-children");
            if (childrenEl && hasChildren) {
                childrenEl.classList.add("open");
                toggle.classList.add("expanded");
            }
        });

        // 点击箭头: 展开/折叠
        toggle.addEventListener("click", (e) => {
            e.stopPropagation();
            const childrenEl = wrap.querySelector(":scope > .tree-children");
            if (!childrenEl) return;
            const open = childrenEl.classList.toggle("open");
            toggle.classList.toggle("expanded", open);
        });

        wrap.appendChild(row);

        // 子节点
        if (hasChildren) {
            const childWrap = document.createElement("div");
            childWrap.className = "tree-children";
            for (const c of children) {
                if (c.type !== "folder") continue;
                if (state.folderFilter) {
                    // 过滤模式: 只显示匹配或包含匹配的
                    if (!folderMatchesFilter(c)) continue;
                }
                childWrap.appendChild(createTreeNode(c, depth + 1, false));
            }
            wrap.appendChild(childWrap);
            // 过滤模式下默认展开
            if (state.folderFilter) {
                childWrap.classList.add("open");
                toggle.classList.add("expanded");
            }
        }
        return wrap;
    }

    function folderMatchesFilter(folder) {
        const kw = state.folderFilter.toLowerCase();
        if (!kw) return true;
        if ((folder.name || "").toLowerCase().includes(kw)) return true;
        // 任意子孙文件夹匹配也算
        for (const c of (folder.children || [])) {
            if (c.type === "folder" && folderMatchesFilter(c)) return true;
        }
        return false;
    }

    function updateSidebarFooter() {
        const stats = state.data.stats || {};
        dom.sidebarFooter.innerHTML = `
            <div>目录: <strong>${stats.folders || 0}</strong>　文件: <strong>${stats.files || 0}</strong></div>
            <div>缩略图: <strong>${stats.thumbnails || 0}</strong>　生成于 ${escapeHtml(state.data.generated_at || "")}</div>
        `;
    }

    // ============ 选中文件夹 ============
    function selectFolder(folderId) {
        state.currentFolderId = folderId || "";
        // 高亮: 清除所有, 再选中目标
        document.querySelectorAll(".tree-row").forEach(r => r.classList.remove("active"));
        if (!folderId) {
            // 根节点: 选中第一个 tree-row (即"全部"节点)
            const rootRow = dom.treeView.querySelector(".tree-node > .tree-row");
            if (rootRow) rootRow.classList.add("active");
        } else {
            const target = document.querySelector(`.tree-row[data-folder-id="${CSS.escape(folderId)}"]`);
            if (target) {
                target.classList.add("active");
                target.scrollIntoView({ block: "nearest", behavior: "smooth" });
            }
        }
        // 重置滚动位置
        dom.cardGrid.scrollTop = 0;
        dom.backToTop.classList.add("hide");
        renderBreadcrumb();
        renderCards();
    }

    function renderBreadcrumb() {
        dom.breadcrumb.innerHTML = "";
        const chain = [];
        let cur = state.currentFolderId;
        while (cur) {
            const n = state.folderMap.get(cur);
            if (!n) break;
            chain.unshift(n);
            cur = n.pid || "";
        }
        // 根节点
        const root = document.createElement("span");
        root.className = "crumb";
        root.textContent = "全部";
        root.dataset.folderId = "";
        root.addEventListener("click", () => selectFolder(""));
        if (chain.length === 0) root.classList.add("current");
        dom.breadcrumb.appendChild(root);

        chain.forEach((n, i) => {
            const sep = document.createElement("span");
            sep.className = "crumb-sep";
            sep.textContent = "/";
            dom.breadcrumb.appendChild(sep);

            const c = document.createElement("span");
            c.className = "crumb" + (i === chain.length - 1 ? " current" : "");
            c.textContent = n.name;
            c.title = n.path || n.name;
            c.dataset.folderId = n.id;
            c.addEventListener("click", () => selectFolder(n.id));
            dom.breadcrumb.appendChild(c);
        });
    }

    // ============ 获取当前文件列表 ============
    function getCurrentFiles() {
        const extFilter = state.extFilter;
        const applyExt = (f) => !extFilter || (f.ext || "").toLowerCase() === extFilter;

        // 搜索模式: 全局检索(可同时应用类型筛选)
        if (state.keyword) {
            const all = state.data.files || [];
            const kw = state.keyword.toLowerCase();
            return all.filter(f =>
                applyExt(f) &&
                (
                    (f.name || "").toLowerCase().includes(kw) ||
                    (f.rel_path || "").toLowerCase().includes(kw) ||
                    (f.ext || "").toLowerCase().includes(kw)
                )
            );
        }
        // 根节点("全部"): 显示所有文件, 按类型筛选
        if (!state.currentFolderId) {
            const all = state.data.files || [];
            return extFilter ? all.filter(applyExt) : all;
        }
        // 具体目录: 先显示本目录直接文件, 再递归显示所有子目录文件
        const folder = state.folderMap.get(state.currentFolderId);
        if (!folder) return [];
        const files = [];
        // 递归收集: 先本目录文件, 再子目录文件
        function collectFiles(node) {
            if (!node) return;
            const children = node.children || [];
            // 先收集本层文件
            for (const c of children) {
                if (c && c.type !== "folder") files.push(c);
            }
            // 再递归子目录
            for (const c of children) {
                if (c && c.type === "folder") collectFiles(c);
            }
        }
        collectFiles(folder);
        return extFilter ? files.filter(applyExt) : files;
    }

    // ============ 渲染后缀筛选下拉框 ============
    function renderExtFilter() {
        const sel = dom.extFilter;
        const exts = (state.data && state.data.exts) || [];
        // 保留当前选择
        const prev = state.extFilter;
        sel.innerHTML = "";
        // "全部类型" 选项
        const allOpt = document.createElement("option");
        allOpt.value = "";
        allOpt.textContent = exts.length ? `全部类型 (${state.data.files.length})` : "全部类型";
        sel.appendChild(allOpt);
        // 实际存在的后缀
        for (const e of exts) {
            const opt = document.createElement("option");
            opt.value = e.ext;
            opt.textContent = `.${e.ext}  (${e.count})`;
            sel.appendChild(opt);
        }
        // 恢复选择, 如已不存在则重置
        if (prev && exts.some(e => e.ext === prev)) {
            sel.value = prev;
        } else {
            sel.value = "";
            state.extFilter = "";
        }
        // 没有任何后缀时禁用
        sel.disabled = exts.length === 0;
        sel.title = exts.length === 0
            ? "尚未扫描到任何文档"
            : `共 ${exts.length} 种类型, 按文件数排序`;
    }

    // ============ 渲染卡片 ============
    function renderCards() {
        let files = getCurrentFiles();
        files = sortFiles(files);

        dom.cardGrid.innerHTML = "";
        dom.cardGrid.className = "card-grid view-" + state.viewMode;

        // 文件夹信息
        const folder = state.folderMap.get(state.currentFolderId);
        const folderName = folder ? (folder.name || "全部") : "全部";
        let info;
        if (state.keyword && state.extFilter) {
            info = `全局检索: "${state.keyword}" · 类型: .${state.extFilter}`;
        } else if (state.keyword) {
            info = `全局检索: "${state.keyword}"`;
        } else if (state.extFilter) {
            info = `${folderName} · 类型: .${state.extFilter}`;
        } else {
            info = `当前目录: ${folderName}`;
        }
        dom.folderInfo.textContent = info;

        dom.resultCount.textContent = `${files.length} 项`;

        if (files.length === 0) {
            dom.emptyState.style.display = "block";
            return;
        }
        dom.emptyState.style.display = "none";

        // 一次性渲染全部卡片
        const frag = document.createDocumentFragment();
        for (const f of files) {
            frag.appendChild(createCard(f));
        }
        dom.cardGrid.appendChild(frag);
    }

    function createCard(file) {
        const card = document.createElement("div");
        card.className = "card";
        card.dataset.fileId = file.id;
        card.title = file.name + "\n" + (file.path || "") + "\n" + (file.size_text || "");

        // 缩略图
        const thumb = document.createElement("div");
        thumb.className = "card-thumb";
        const img = document.createElement("img");
        img.loading = "lazy";
        img.alt = file.name;
        const thumbUrl = file.thumbnail ? (file.thumbnail + "?t=" + (state.data.generated_at || "").replace(/\W/g, "")) : "";
        img.src = thumbUrl || "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCA1IDciLz4=";
        img.onerror = () => {
            img.onerror = null;
            img.src = makePlaceholderSvg(file);
        };
        thumb.appendChild(img);

        // 类型徽标
        const badge = document.createElement("span");
        badge.className = "ext-badge";
        badge.textContent = (file.ext || "?").toUpperCase();
        thumb.appendChild(badge);

        // 内容
        const body = document.createElement("div");
        body.className = "card-body";
        const title = document.createElement("div");
        title.className = "card-title";
        title.textContent = file.name;
        const meta = document.createElement("div");
        meta.className = "card-meta";
        const size = document.createElement("span");
        size.className = "size";
        size.textContent = file.size_text || "";
        const path = document.createElement("span");
        path.className = "path-hint";
        path.textContent = file.rel_path || "";
        path.style.color = "var(--text-soft)";
        path.style.fontSize = "11px";
        path.style.overflow = "hidden";
        path.style.textOverflow = "ellipsis";
        path.style.whiteSpace = "nowrap";
        path.style.maxWidth = "100px";
        meta.appendChild(size);
        meta.appendChild(path);
        body.appendChild(title);
        body.appendChild(meta);

        // 卡片操作按钮
        const actions = document.createElement("div");
        actions.className = "card-actions";
        const btnOpen = document.createElement("button");
        btnOpen.title = "在浏览器打开";
        btnOpen.textContent = "🌐";
        btnOpen.addEventListener("click", (e) => { e.stopPropagation(); openFileInBrowser(file); });
        const btnPreview = document.createElement("button");
        btnPreview.title = "预览";
        btnPreview.textContent = "🔍";
        btnPreview.addEventListener("click", (e) => { e.stopPropagation(); openPreview(file); });
        actions.appendChild(btnPreview);
        actions.appendChild(btnOpen);

        card.appendChild(thumb);
        card.appendChild(body);
        card.appendChild(actions);

        // 卡片点击: 预览
        card.addEventListener("click", () => openPreview(file));
        // 右键: 菜单
        card.addEventListener("contextmenu", (e) => {
            e.preventDefault();
            showCardMenu(e.clientX, e.clientY, file);
        });

        return card;
    }

    function makePlaceholderSvg(file) {
        const ext = (file.ext || "?").toUpperCase();
        const colors = {
            PDF: ["#dc2626", "#fca5a5"],
            DOC: ["#2563eb", "#93c5fd"], DOCX: ["#2563eb", "#93c5fd"],
            TXT: ["#059669", "#86efac"], MD: ["#7c3aed", "#c4b5fd"],
            RTF: ["#d97706", "#fcd34d"],
        };
        const [bg, fg] = colors[ext] || ["#6b7280", "#d1d5db"];
        const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 5 7">
            <rect width="5" height="7" fill="${bg}"/>
            <text x="2.5" y="4" font-size="2" fill="#fff" text-anchor="middle" font-family="sans-serif" font-weight="bold">${ext}</text>
        </svg>`;
        return "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(svg)));
    }

    // ============ 排序 ============
    function sortFiles(files) {
        const arr = files.slice();
        const mode = state.sortBy;
        if (mode === "natural") {
            arr.sort((a, b) => naturalCompare(a.name, b.name));
        } else if (mode === "name-asc") {
            arr.sort((a, b) => naturalCompare(a.name, b.name));
        } else if (mode === "name-desc") {
            arr.sort((a, b) => -naturalCompare(a.name, b.name));
        } else if (mode === "size-asc") {
            arr.sort((a, b) => (a.size || 0) - (b.size || 0));
        } else if (mode === "size-desc") {
            arr.sort((a, b) => (b.size || 0) - (a.size || 0));
        } else if (mode === "ext-asc") {
            arr.sort((a, b) => (a.ext || "").localeCompare(b.ext || "") || naturalCompare(a.name, b.name));
        } else if (mode === "ext-desc") {
            arr.sort((a, b) => (b.ext || "").localeCompare(a.ext || "") || naturalCompare(a.name, b.name));
        }
        return arr;
    }

    // ============ 预览弹窗 (画册模式) ============
    function openPreview(file) {
        currentPreviewFile = file;
        dom.previewTitle.textContent = file.name;

        // 收集所有缩略图 (支持 thumbnails 数组, 兼容旧数据只有单张)
        galleryImages = (file.thumbnails && file.thumbnails.length > 0)
            ? [...file.thumbnails]
            : (file.thumbnail ? [file.thumbnail] : []);
        galleryIndex = 0;

        _showGalleryImage();
        dom.previewMeta.innerHTML = `
            <div class="meta-row"><span class="k">文件类型</span><span class="v">.${escapeHtml(file.ext || "")} (${escapeHtml(file.category || "")})</span></div>
            <div class="meta-row"><span class="k">文件大小</span><span class="v">${escapeHtml(file.size_text || "")}</span></div>
            <div class="meta-row"><span class="k">相对路径</span><span class="v">${escapeHtml(file.rel_path || "")}</span></div>
            <div class="meta-row"><span class="k">完整路径</span><span class="v" style="font-family:monospace;font-size:12px">${escapeHtml(file.path || "")}</span></div>
            <div class="meta-row"><span class="k">所在目录</span><span class="v" style="font-family:monospace;font-size:12px">${escapeHtml(parentDirOf(file.path || ""))}</span></div>
        `;
        dom.previewModal.style.display = "flex";
    }

    // 显示画册当前图片
    function _showGalleryImage() {
        const file = currentPreviewFile;
        if (galleryImages.length === 0) {
            dom.previewImg.src = makePlaceholderSvg(file);
        } else {
            dom.previewImg.src = galleryImages[galleryIndex];
            dom.previewImg.onerror = () => {
                dom.previewImg.onerror = null;
                dom.previewImg.src = makePlaceholderSvg(file);
            };
        }
        // 更新导航按钮和页码
        const hasMultiple = galleryImages.length > 1;
        dom.galleryPrev.style.display = hasMultiple ? "flex" : "none";
        dom.galleryNext.style.display = hasMultiple ? "flex" : "none";
        if (hasMultiple) {
            dom.galleryCounter.style.display = "block";
            dom.galleryCounter.textContent = `${galleryIndex + 1} / ${galleryImages.length}`;
        } else {
            dom.galleryCounter.style.display = "none";
        }
    }

    function galleryPrev() {
        if (galleryImages.length <= 1) return;
        galleryIndex = (galleryIndex - 1 + galleryImages.length) % galleryImages.length;
        _showGalleryImage();
    }

    function galleryNext() {
        if (galleryImages.length <= 1) return;
        galleryIndex = (galleryIndex + 1) % galleryImages.length;
        _showGalleryImage();
    }

    // 双击放大
    function openLightbox(src) {
        dom.lightboxImg.src = src;
        dom.imageLightbox.style.display = "flex";
    }
    function closeLightbox() {
        dom.imageLightbox.style.display = "none";
    }

    function closePreview() {
        dom.previewModal.style.display = "none";
        currentPreviewFile = null;
        galleryImages = [];
        galleryIndex = 0;
        closeLightbox();
    }

    // ============ 本地操作 API ============
    async function apiPost(url, body) {
        try {
            const res = await fetch(url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(body),
            });
            return await res.json();
        } catch (e) {
            return { ok: false, error: "无法连接服务器: " + e.message };
        }
    }

    function openFileInBrowser(file) {
        const url = fileUriFromPath(file.path);
        showToast("正在打开文件...", 1500);
        window.open(url, "_blank");
    }

	function openFileServe(file) {
	    const url = fileUriFromPath(file.path);
	    showToast("正在尝试打开文件...", 1500);
	    window.open(url, "_blank");
	}

    async function tryOpenDir(file) {
        showToast("正在打开文件位置...", 1500);
        const r = await apiPost("/api/open-folder", { path: file.path });
        if (r.ok) {
            showToast("已打开文件位置", 1500);
        } else {
            showToast("打开失败: " + (r.error || "未知错误"), 3000);
        }
    }

    // ============ 右键菜单 ============
    function showCardMenu(x, y, file) {
        dom.cardMenu.style.display = "block";
        dom.cardMenu.style.left = x + "px";
        dom.cardMenu.style.top = y + "px";
        dom.cardMenu.dataset.fileId = file.id;
        // 边界处理
        const rect = dom.cardMenu.getBoundingClientRect();
        if (rect.right > window.innerWidth) {
            dom.cardMenu.style.left = (window.innerWidth - rect.width - 10) + "px";
        }
        if (rect.bottom > window.innerHeight) {
            dom.cardMenu.style.top = (window.innerHeight - rect.height - 10) + "px";
        }
    }
    function hideCardMenu() {
        dom.cardMenu.style.display = "none";
    }

    // ============ 事件绑定 ============
    function bindEvents() {
        // 视图切换
        dom.viewGrid.addEventListener("click", () => setViewMode("grid"));
        dom.viewList.addEventListener("click", () => setViewMode("list"));
        // 返回顶部: 滚动判断显示（滚动超过阈值才显示，否则隐藏）
        const THRESHOLD = 300;
        let scrollTimer;
        function updateBackTopBtn() {
            const shouldHide = dom.cardGrid.scrollTop <= THRESHOLD;
            dom.backToTop.classList.toggle("hide", shouldHide);
        }
        // 立即检查一次（切换目录或搜索后立即应用正确可见性）
        updateBackTopBtn();
        dom.cardGrid.addEventListener("scroll", () => {
            if (scrollTimer) return;
            scrollTimer = setTimeout(() => { scrollTimer = null; updateBackTopBtn(); }, 80);
        });
        // 返回顶部: 点击平滑滚回
        dom.backToTop.addEventListener("click", () => {
            dom.cardGrid.scrollTo({ top: 0, behavior: "smooth" });
        });
        // 排序
        dom.sortSelect.addEventListener("change", () => {
            state.sortBy = dom.sortSelect.value;
            renderCards();
        });
        // 后缀筛选
        dom.extFilter.addEventListener("change", () => {
            state.extFilter = dom.extFilter.value;
            renderCards();
        });
        // 搜索
        let searchTimer;
        dom.searchInput.addEventListener("input", () => {
            clearTimeout(searchTimer);
            searchTimer = setTimeout(() => {
                state.keyword = dom.searchInput.value.trim();
                dom.searchClear.style.display = state.keyword ? "flex" : "none";
                renderCards();
            }, 200);
        });
        dom.searchClear.addEventListener("click", () => {
            dom.searchInput.value = "";
            state.keyword = "";
            dom.searchClear.style.display = "none";
            renderCards();
        });
        // 目录过滤
        let folderTimer;
        dom.folderFilter.addEventListener("input", () => {
            clearTimeout(folderTimer);
            folderTimer = setTimeout(() => {
                state.folderFilter = dom.folderFilter.value.trim();
                renderTree();
            }, 200);
        });
        // 展开/折叠全部
        dom.expandAll.addEventListener("click", () => {
            document.querySelectorAll(".tree-children").forEach(el => el.classList.add("open"));
            document.querySelectorAll(".tree-toggle").forEach(el => el.classList.add("expanded"));
        });
        dom.collapseAll.addEventListener("click", () => {
            // 保留第一层展开
            const first = document.querySelector(".tree-node > .tree-children");
            document.querySelectorAll(".tree-children").forEach(el => {
                if (el !== first) el.classList.remove("open");
            });
            document.querySelectorAll(".tree-toggle").forEach(el => {
                // 简单处理: 全部移除, 再把第一个加回
                el.classList.remove("expanded");
            });
            if (first) {
                first.classList.add("open");
                const t = first.parentElement.querySelector(":scope > .tree-row > .tree-toggle");
                if (t) t.classList.add("expanded");
            }
        });
        // 刷新
        dom.refreshBtn.addEventListener("click", () => location.reload());
        // 重新扫描
        dom.rescanBtn.addEventListener("click", () => {
            showToast("请运行 scanner/dist/DocumentScanner.exe 重新扫描", 4000);
            // 提示打开方式
            const ok = confirm("需要运行扫描器重新生成数据.\n\n点击确定查看说明, 取消则关闭.");
            if (ok) {
                alert("使用方法:\n1. 双击 scanner/dist/DocumentScanner.exe\n2. 选择主目录\n3. 等待扫描完成\n4. 回到此页面点击 ↻ 刷新");
            }
        });
        // 关闭弹窗
        dom.modalClose.addEventListener("click", closePreview);
        dom.previewModal.querySelector(".modal-backdrop").addEventListener("click", closePreview);
        // 画册导航
        dom.galleryPrev.addEventListener("click", (e) => { e.stopPropagation(); galleryPrev(); });
        dom.galleryNext.addEventListener("click", (e) => { e.stopPropagation(); galleryNext(); });
        // 双击图片放大
        dom.previewImg.addEventListener("dblclick", () => {
            if (galleryImages.length > 0) {
                openLightbox(galleryImages[galleryIndex]);
            }
        });
        // lightbox 关闭
        dom.lightboxClose.addEventListener("click", closeLightbox);
        dom.imageLightbox.querySelector(".lightbox-backdrop").addEventListener("click", closeLightbox);
        // 键盘: ESC 关闭, 左右切换
        document.addEventListener("keydown", (e) => {
            if (e.key === "Escape" || e.code === "Escape" || e.keyCode === 27) {
                if (dom.imageLightbox.style.display === "flex") {
                    closeLightbox();
                } else {
                    closePreview();
                    hideCardMenu();
                }
            }
            // 预览弹窗中左右切换
            if (dom.previewModal.style.display === "flex" && galleryImages.length > 1) {
                if (e.key === "ArrowLeft" || e.keyCode === 37) galleryPrev();
                if (e.key === "ArrowRight" || e.keyCode === 39) galleryNext();
            }
        });
        // 弹窗操作按钮
        dom.openInBrowserBtn.addEventListener("click", () => {
            if (currentPreviewFile) openFileServe(currentPreviewFile);
        });
        dom.openFileBtn.addEventListener("click", () => {
            if (currentPreviewFile) openFileInBrowser(currentPreviewFile);
        });
        dom.copyPathBtn.addEventListener("click", () => {
            if (currentPreviewFile) copyText(currentPreviewFile.path);
        });
        dom.openParentBtn.addEventListener("click", () => {
            if (currentPreviewFile) tryOpenDir(currentPreviewFile);
        });
        // 右键菜单事件
        dom.cardMenu.addEventListener("click", (e) => {
            const btn = e.target.closest("button");
            if (!btn) return;
            const act = btn.dataset.act;
            const fileId = dom.cardMenu.dataset.fileId;
            const file = (state.data.files || []).find(f => f.id === fileId);
            hideCardMenu();
            if (!file) return;
            if (act === "preview") openPreview(file);
            else if (act === "openInBrowser") openFileServe(file);
            else if (act === "open") openFileInBrowser(file);
            else if (act === "copyPath") copyText(file.path);
            else if (act === "tryOpenDir") tryOpenDir(file);
        });
        // 点击其他位置关闭右键菜单
        document.addEventListener("click", () => hideCardMenu());
        document.addEventListener("scroll", () => hideCardMenu(), true);
        // 重试按钮
        dom.retryLoadBtn.addEventListener("click", () => location.reload());
    }

    function setViewMode(mode) {
        state.viewMode = mode;
        dom.viewGrid.classList.toggle("active", mode === "grid");
        dom.viewList.classList.toggle("active", mode === "list");
        renderCards();
    }

    // ============ 初始化 ============
    async function init() {
        dom.loadingText.textContent = "正在加载数据…";
        try {
            state.data = await loadData();
        } catch (e) {
            console.error("加载数据失败:", e);
            dom.loadingMask.style.display = "none";
            dom.firstRunTip.style.display = "flex";
            return;
        }
        if (!state.data || !state.data.tree) {
            dom.loadingMask.style.display = "none";
            dom.firstRunTip.style.display = "flex";
            return;
        }
        buildIndex();
        renderExtFilter();
        renderTree();
        renderBreadcrumb();
        renderCards();
        bindEvents();
        dom.loadingMask.style.display = "none";
    }

    document.addEventListener("DOMContentLoaded", init);
})();
