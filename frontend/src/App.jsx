import { useState } from 'react'
import axios from 'axios'

function App() {
    const [file, setFile] = useState(null)
    const [status, setStatus] = useState("")
    const [downloadUrl, setDownloadUrl] = useState("")

    const handleDrop = (e) => {
        e.preventDefault()
        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            setFile(e.dataTransfer.files[0])
            setStatus("File selected: " + e.dataTransfer.files[0].name)
        }
    }

    const handleUpload = async () => {
        if (!file) return;
        setStatus("Uploading and Processing...")
        const formData = new FormData()
        formData.append("file", file)

        try {
            const res = await axios.post('/upload', formData)
            setStatus("Processing Complete!")
            setDownloadUrl(res.data.download_url)
        } catch (err) {
            console.error(err)
            setStatus("Error: " + (err.response?.data?.error || err.message))
        }
    }

    return (
        <div className="min-h-screen bg-gray-100 flex items-center justify-center p-4">
            <div className="bg-white p-8 rounded-xl shadow-lg w-full max-w-md">
                <h1 className="text-2xl font-bold mb-6 text-center text-gray-800">PPT Generator</h1>

                <div
                    className="border-dashed border-4 border-gray-300 rounded-lg p-10 text-center cursor-pointer hover:bg-gray-50 transition"
                    onDragOver={(e) => e.preventDefault()}
                    onDrop={handleDrop}
                    onClick={() => document.getElementById('fileInput').click()}
                >
                    <input
                        id="fileInput"
                        type="file"
                        className="hidden"
                        accept=".csv"
                        onChange={(e) => {
                            if (e.target.files[0]) {
                                setFile(e.target.files[0])
                                setStatus("File selected: " + e.target.files[0].name)
                            }
                        }}
                    />
                    {file ? (
                        <p className="text-green-600 font-medium">{file.name}</p>
                    ) : (
                        <p className="text-gray-500">Drag & Drop CSV here <br /> or Click to Browse</p>
                    )}
                </div>

                <button
                    onClick={handleUpload}
                    disabled={!file}
                    className="w-full mt-6 bg-blue-600 hover:bg-blue-700 text-white font-bold py-3 rounded-lg disabled:opacity-50 transition"
                >
                    Generate Report
                </button>

                {status && <div className="mt-4 text-center text-sm font-semibold text-gray-700">{status}</div>}

                {downloadUrl && (
                    <a
                        href={downloadUrl}
                        download
                        className="block mt-4 text-center text-blue-600 underline"
                    >
                        Download Final Report
                    </a>
                )}
            </div>
        </div>
    )
}

export default App
