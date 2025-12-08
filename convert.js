// Format Converter Script
document.addEventListener("DOMContentLoaded", () => {
    const fileInput = document.getElementById("convFile")
    const statusEl = document.getElementById("convStatus")
    const btn = document.getElementById("btnConvert")
    const typeGrid = document.getElementById("convTypeGrid")
    const progressBar = document.getElementById("convProgress")

    let selectedType = null

    // Dosya tipine göre uygun conversion seçeneklerini göster
    function updateConversionOptions(fileName) {
        if (!fileName) {
            // Dosya seçilmediyse tüm seçenekleri göster ama pasif yap
            typeGrid.querySelectorAll(".convert-type-option").forEach((option) => {
                option.style.opacity = "0.5"
                option.style.pointerEvents = "none"
                option.classList.remove("selected")
                option.querySelector("input").checked = false
            })
            selectedType = null
            return
        }

        const ext = fileName.toLowerCase().split('.').pop()
        const availableTypes = []

        // Dosya tipine göre uygun conversion'ları belirle
        if (ext === 'pptx' || ext === 'ppt') {
            availableTypes.push('pptx-to-pdf', 'pptx-to-word')
        } else if (ext === 'pdf') {
            availableTypes.push('pdf-to-pptx', 'pdf-to-word')
        } else if (ext === 'docx' || ext === 'doc') {
            availableTypes.push('word-to-pdf', 'word-to-pptx')
        }

        // Tüm seçenekleri güncelle
        typeGrid.querySelectorAll(".convert-type-option").forEach((option) => {
            const type = option.dataset.value
            if (availableTypes.includes(type)) {
                option.style.opacity = "1"
                option.style.pointerEvents = "auto"
            } else {
                option.style.opacity = "0.3"
                option.style.pointerEvents = "none"
                if (option.classList.contains("selected")) {
                    option.classList.remove("selected")
                    option.querySelector("input").checked = false
                    if (selectedType === type) {
                        selectedType = null
                    }
                }
            }
        })
    }

    // Dosya seçildiğinde
    fileInput.addEventListener("change", (e) => {
        const file = e.target.files[0]
        if (file) {
            updateConversionOptions(file.name)
            showStatus(`File selected: ${file.name}`, "info")
        } else {
            updateConversionOptions(null)
        }
    })

    // Handle conversion type selection
    typeGrid.querySelectorAll(".convert-type-option").forEach((option) => {
        option.addEventListener("click", () => {
            if (option.style.pointerEvents === "none") return
            
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

        showStatus("Preparing file...", "")
        showProgress(10)
        btn.disabled = true
        btn.innerHTML = `
            <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" class="spin">
                <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
            </svg>
            Converting...
        `

        try {
            showProgress(20)
            showStatus("Reading file...", "")
            const buffer = await file.arrayBuffer()
            const uint8 = new Uint8Array(buffer)
            let binary = ""
            for (let i = 0; i < uint8.length; i++) binary += String.fromCharCode(uint8[i])
            const base64Data = btoa(binary)

            showProgress(30)
            showStatus("Uploading to server...", "")
            const payload = {
                fileName: file.name,
                fileData: base64Data,
                type: selectedType,
            }

            // XMLHttpRequest kullan progress için
            const xhr = new XMLHttpRequest()
            const CONVERT_SERVER = window.CONVERT_SERVER_URL || 'http://localhost:3002';
            xhr.open("POST", `${CONVERT_SERVER}/convert`, true)
            xhr.setRequestHeader("Content-Type", "application/json")
            
            // Progress tracking
            xhr.upload.onprogress = (e) => {
                if (e.lengthComputable) {
                    const uploadPercent = 30 + (e.loaded / e.total) * 20 // 30-50%
                    showProgress(uploadPercent)
                }
            }

            xhr.onprogress = (e) => {
                if (e.lengthComputable) {
                    const downloadPercent = 50 + (e.loaded / e.total) * 40 // 50-90%
                    showProgress(downloadPercent)
                }
            }

            const responsePromise = new Promise((resolve, reject) => {
                xhr.onload = () => {
                    if (xhr.status >= 200 && xhr.status < 300) {
                        resolve(xhr.response)
                    } else {
                        let errorMsg = "Server error"
                        try {
                            const errorJson = JSON.parse(xhr.responseText)
                            errorMsg = errorJson.error || errorMsg
                        } catch (e) {
                            errorMsg = xhr.statusText || errorMsg
                        }
                        reject(new Error(errorMsg))
                    }
                }
                xhr.onerror = () => reject(new Error("Network error"))
                xhr.ontimeout = () => reject(new Error("Request timeout"))
                xhr.timeout = 180000 // 3 dakika timeout
            })

            xhr.send(JSON.stringify(payload))
            showProgress(40)
            showStatus("Converting... This may take a while...", "")

            const arrayBuf = await responsePromise
            showProgress(90)
            showStatus("Downloading converted file...", "")

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

            showProgress(100)
            showStatus("Conversion complete! File downloaded.", "success")
            setTimeout(() => {
                showProgress(0)
            }, 2000)
        } catch (err) {
            console.error(err)
            showProgress(0)
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
        else if (type === "info") statusEl.classList.add("info")
    }

    function showProgress(percent) {
        const progressContainer = document.getElementById("convProgressContainer")
        if (!progressBar || !progressContainer) return
        progressBar.style.width = percent + "%"
        if (percent > 0 && percent < 100) {
            progressContainer.style.display = "block"
        } else if (percent === 100) {
            setTimeout(() => {
                progressContainer.style.display = "none"
            }, 1000)
        } else {
            progressContainer.style.display = "none"
        }
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
