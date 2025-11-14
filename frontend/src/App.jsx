import React, { useState } from "react";
import axios from "axios";
import "./App.css";

function App() {
    const [coords, setCoords] = useState(
        Array.from({ length: 10 }, () => ({ x: "", y: "" }))
    );
    const [description, setDescription] = useState("");
    const [excelFile, setExcelFile] = useState(null);
    const [pdfFile, setPdfFile] = useState(null);
    const [loading, setLoading] = useState(false);

    const handleCoordChange = (index, field, value) => {
        setCoords(prev =>
            prev.map((coord, i) =>
                i === index ? { ...coord, [field]: value } : coord
            )
        );
    };

    const handleFileUpload = setter => e => setter(e.target.files[0]);

    const buildFormData = () => {
        const formData = new FormData();
        formData.append("description", description);

        if (excelFile) {
            formData.append("excel_file", excelFile);
        } else {
            const validCoords = coords
                .filter(c => c.x && c.y)
                .map(c => [parseFloat(c.x), parseFloat(c.y)]);

            formData.append("coords_json", JSON.stringify(validCoords));
        }

        if (pdfFile) {
            formData.append("pdf_file", pdfFile);
        }

        return formData;
    };

    const handleSubmit = async () => {
        setLoading(true);

        try {
    const formData = buildFormData();

    const response = await axios.post(
        "http://localhost:8000/generate",
        formData,
        { responseType: "blob" }
    );

    const url = window.URL.createObjectURL(new Blob([response.data]));
    const link = document.createElement("a");

    link.href = url;
    link.download = "generated_document.docx";
    link.click();
    link.remove();

} catch (err) {
    console.error("Ошибка:", err);

    let message = "Ошибка при генерации документа";

    if (!err.response) {
            message = "Сервер недоступен";
    } else {
        try {
            const data = err.response.data;

            const text =
                typeof data === "string"
                    ? data
                    : await data.text();

            try {
                const json = JSON.parse(text);

                if (json.detail?.error) {
                    message = json.detail.error;
                } else if (json.error) {
                    message = json.error;
                }
            } catch (parseErr) {
                console.warn("Ошибка парсинга JSON:", parseErr);
            }
        } catch (ex) {
            console.warn("Ошибка извлечения текста ошибки:", ex);
        }
    }

        alert(message);
    } finally {
        setLoading(false);
    }
    };

    return (
        <div className="container">
            <h2>Генератор отчетов</h2>

            <div className="section">
                <label className="label">Описание:</label>
                <textarea
                    className="textareaField"
                    rows={4}
                    value={description}
                    onChange={e => setDescription(e.target.value)}
                />
            </div>

            <div className="section">
                <h3>Координаты (x/y)</h3>
                {coords.map((coord, i) => (
                    <div key={i} className="coordRow">
                        <input
                            type="number"
                            placeholder="x"
                            className="inputField"
                            value={coord.x}
                            onChange={e => handleCoordChange(i, "x", e.target.value)}
                        />
                        <input
                            type="number"
                            placeholder="y"
                            className="inputField"
                            value={coord.y}
                            onChange={e => handleCoordChange(i, "y", e.target.value)}
                        />
                    </div>
                ))}
            </div>

            <div className="section">
                <label className="label">Или загрузить Excel с координатами:</label>
                <input
                    type="file"
                    accept=".xlsx,.xls"
                    onChange={handleFileUpload(setExcelFile)}
                />
            </div>

            <div className="section">
                <label className="label">Загрузить PDF:</label>
                <input
                    type="file"
                    accept=".pdf"
                    onChange={handleFileUpload(setPdfFile)}
                />
            </div>

            <div className="center">
                <button
                    className={`button ${loading ? "buttonDisabled" : ""}`}
                    onClick={handleSubmit}
                    disabled={loading}
                >
                    {loading && <span className="loader" />}
                    {loading ? "Генерация..." : "Сгенерировать DOCX"}
                </button>
            </div>
        </div>
    );
}

export default App;
