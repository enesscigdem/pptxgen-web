document.getElementById("fileInput").addEventListener("change", async e => {

    let file = e.target.files[0];
    if (!file) return;

    let loading = document.getElementById("loading");
    let out = document.getElementById("output");
    out.innerHTML = "";
    loading.style.display = "block";

    let arrayBuffer = await file.arrayBuffer();
    let zip = await JSZip.loadAsync(arrayBuffer);

    let parser = new DOMParser();

    let presXml = await zip.file("ppt/presentation.xml").async("text");
    let presDoc = parser.parseFromString(presXml, "application/xml");
    let sldSz = presDoc.querySelector("p\\:sldSz,sldSz");
    let slideW = sldSz.getAttribute("cx");
    let slideH = sldSz.getAttribute("cy");

    let slidePaths = Object.keys(zip.files)
        .filter(p => /ppt\/slides\/slide\d+\.xml$/.test(p))
        .sort((a, b) => parseInt(a.match(/slide(\d+)/)[1]) - parseInt(b.match(/slide(\d+)/)[1]));

    let fullJsonOutput = [];

    for (let p of slidePaths) {

        let xml = await zip.file(p).async("text");
        let doc = parser.parseFromString(xml, "application/xml");

        let slideJson = {
            slideNumber: parseInt(p.match(/slide(\d+)/)[1]),
            slideWidth: +slideW,
            slideHeight: +slideH,
            elements: []
        };

        let index = 0;

        let shapes = doc.querySelectorAll("p\\:sp,sp");
        shapes.forEach(sp => {
            let txtNodes = [...sp.querySelectorAll("a\\:t,t")].map(t => t.textContent.trim()).filter(Boolean);
            if (!txtNodes.length) return;

            let text = txtNodes.join(" ");

            let elem = { type: "textbox", text, zIndex: ++index };

            let ph = sp.querySelector("p\\:ph,ph");
            if (ph) {
                let type = ph.getAttribute("type");
                if (type === "title" || type === "ctrTitle") elem.type = "title";
                if (type === "body") elem.type = "content";
                if (type === "sldNum") elem.type = "slideNumber";
                if (type === "dt") elem.type = "date";
            }

            let xfrm = sp.querySelector("a\\:xfrm,xfrm");
            if (xfrm) {
                let off = xfrm.querySelector("a\\:off,off");
                let ext = xfrm.querySelector("a\\:ext,ext");
                elem.position = {
                    x: +off.getAttribute("x"),
                    y: +off.getAttribute("y"),
                    w: +ext.getAttribute("cx"),
                    h: +ext.getAttribute("cy")
                };
            }

            let rPr = sp.querySelector("a\\:rPr,rPr");
            elem.style = {
                size: rPr?.getAttribute("sz") ? parseInt(rPr.getAttribute("sz")) / 100 : null,
                bold: rPr?.getAttribute("b") === "1",
                italic: rPr?.getAttribute("i") === "1",
                align: sp.querySelector("a\\:pPr,a\\:algn")?.textContent || null,
                font: null,
                color: null
            };

            let latin = rPr?.querySelector("a\\:latin,latin");
            if (latin) elem.style.font = latin.getAttribute("typeface");

            let solidFill = sp.querySelector("a\\:solidFill,a\\:srgbClr, srgbClr");
            if (solidFill) elem.style.color = "#" + solidFill.getAttribute("val");

            slideJson.elements.push(elem);
        });

        let pics = doc.querySelectorAll("p\\:pic,pic");
        pics.forEach(pic => {
            let elem = { type: "image", zIndex: ++index };
            let xf = pic.querySelector("a\\:xfrm,xfrm");
            if (xf) {
                let off = xf.querySelector("a\\:off,off");
                let ext = xf.querySelector("a\\:ext,ext");
                elem.position = {
                    x: +off.getAttribute("x"),
                    y: +off.getAttribute("y"),
                    w: +ext.getAttribute("cx"),
                    h: +ext.getAttribute("cy")
                };
            }
            slideJson.elements.push(elem);
        });

        fullJsonOutput.push(slideJson);

        let card = document.createElement("div");
        card.className = "card";

        card.innerHTML = `
            <div class="flex justify-between items-center mb-2">
                <h2 class="font-bold text-lg">📌 Slayt ${slideJson.slideNumber}</h2>
                <button class="copyBtn text-xs bg-gray-100 px-2 py-1 rounded border">
                    Kopyala
                </button>
            </div>
            <pre>${JSON.stringify(slideJson, null, 2)}</pre>
        `;

        out.appendChild(card);
    }

    loading.style.display = "none";

    // copy button events
    document.querySelectorAll(".copyBtn").forEach(btn=>{
        btn.onclick = ()=>{
            let pre = btn.parentElement.nextElementSibling.innerText;
            navigator.clipboard.writeText(pre);
            btn.innerText = "Kopyalandı ✔";
            setTimeout(()=>btn.innerText="Kopyala", 1000);
        }
    });
});
