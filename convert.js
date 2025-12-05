// Format Converter Script
document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("convFile")
    const statusEl = document.getElementById("convStatus")
    const btn = document.getElementById("btnConvert")
    const typeGrid = document.getElementById("convTypeGrid")

    let selectedType = null

    // Handle conversion type selection
    typeGrid.querySelectorAll(".convert-type-option").forEach((option) => {
        option.addEventListener("click", () => {
            typeGrid.querySelectorAll(".convert-type-option").forEach((o) => o.classList.remove("selected"))
            option.classList.add("selected")
            option.querySelector("input").checked = true
            selectedType = option.dataset.value
        })
    })

    btn.addEventListener("click", async () => {
        const file = fileInput.files[0]

        if (!file) {
            showStatus("Please select a file to convert.", "error")
            return
        }
        if (!selectedType) {
            showStatus("Please select a conversion type.", "error")
            return
        }

        showStatus("Starting conversion...", "")
        btn.disabled = true
        btn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin">
                <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
            </svg>
            Converting...
        `

        try {
            const buffer = await file.arrayBuffer()
            const uint8 = new Uint8Array(buffer)
            let binary = ""
            for (let i = 0; i < uint8.length; i++) binary += String.fromCharCode(uint8[i])
            const base64Data = btoa(binary)

            const payload = {
                fileName: file.name,
                fileData: base64Data,
                type: selectedType,
            }

            const response = await fetch("http://localhost:3000/convert", {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            })

            if (!response.ok) {
                throw new Error("Server error")
            }

            const arrayBuf = await response.arrayBuffer()
            const resBlob = new Blob([arrayBuf])
            let outName = file.name

            const extMap = {
                "pptx-to-pdf": ".pdf",
                "pdf-to-pptx": ".pptx",
                "pptx-to-word": ".docx",
                "word-to-pptx": ".pptx",
                "pdf-to-word": ".docx",
                "word-to-pdf": ".pdf",
            }

            const newExt = extMap[selectedType] || ""
            if (newExt) {
                outName = outName.replace(/\.[^.]+$/, newExt)
            }

            const url = window.URL.createObjectURL(resBlob)
            const a = document.createElement("a")
            a.href = url
            a.download = outName
            document.body.appendChild(a)
            a.click()
            a.remove()
            window.URL.revokeObjectURL(url)

            showStatus("Conversion complete! File downloaded.", "success")
        } catch (err) {
            console.error(err)
            showStatus("Conversion failed: " + err.message, "error")
        } finally {
            btn.disabled = false
            btn.innerHTML = `
                <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2">
                    <polyline points="16,3 21,3 21,8"/>
                    <line x1="4" y1="20" x2="21" y2="3"/>
                </svg>
                Convert & Download
            `
        }
    })

    function showStatus(message, type) {
        statusEl.textContent = message
        statusEl.className = "converter-status visible"
        if (type === "success") statusEl.classList.add("success")
        else if (type === "error") statusEl.classList.add("error")
    }
})

// Add spinning animation
const style = document.createElement("style")
style.textContent = `
    @keyframes spin {
        to { transform: rotate(360deg); }
    }
    .spin {
        animation: spin 1s linear infinite;
    }
`
document.head.appendChild(style)
