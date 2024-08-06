"use client";
import { useEffect, useState } from "react";
import axios from "axios";
import { Workbook } from "@fortune-sheet/react";
import "@fortune-sheet/react/dist/index.css";
export default function Home() {
  const [dynamicData, setDynamicData] = useState(null);
  const handleFileUpload = async (event) => {
    const file = event.target.files[0];
    const formData = new FormData();
    formData.append("file", file);
    const response = await axios.post(
      "http://localhost:3001/api/excel",
      formData,
      {
        headers: {
          "Content-Type": "multipart/form-data",
        },
      }
    );
    console.log("setData", response);
    setDynamicData(response.data);
  };
  const formatDataForFortuneSheets = (data) => {
    return data.map((row) =>
      row.map((cell) => ({
        v: cell.value,
        s: {
          font: cell.style.font,
          fill: cell.style.fill,
          alignment: cell.style.alignment,
          border: cell.style.border,
        },
        mc: cell.mergedCells ? { r: 0, c: 0 } : undefined,
      }))
    );
  };
  console.log("setData", dynamicData);
  return (
    <div className="container mx-auto p-4">
      <input type="file" onChange={handleFileUpload} />
      <div
        style={{
          width: "100%",
          height: "700px",
        }}
      >
        {dynamicData && <Workbook data={[dynamicData]} />}
      </div>
    </div>
  );
}