import { useState } from "react";
import { useNotifyContext } from "../contexts/notifyContext";
import spinner from "../images/loading-part-2-svgrepo-com.svg";

const RunInsert = () => {
  const ipcRenderer = window.ipcRenderer;
  const { show } = useNotifyContext();
  const [isLoading, setIsLoading] = useState(false);

  const handleRun = () => {
    setIsLoading(true);
    ipcRenderer.send("run-insert", {});
    ipcRenderer.once("run-insert", (event, data) => {
      if (data.error) {
        show(data.error, "error");
      } else {
        show("Chèn ảnh thành công", "success");
      }
      setIsLoading(false);
    });
  };

  return (
    <div className="bg-gray-900 w-full h-screen flex flex-col justify-center items-center text-center text-white px-5">
      <p className="mb-4 text-lg font-semibold px-7">
        <span className="text-green-400">Run </span>sau khi đã thêm ảnh vào
        folder <span className="text-green-400">images</span> và nội dung vào
        folder <span className="text-green-400">text</span>
      </p>
      <button
        onClick={handleRun}
        className="bg-gradient-to-r from-green-500 to-green-700 hover:from-green-600 hover:to-green-800 text-white font-semibold rounded-full py-3 px-10 shadow-lg transition duration-300 transform hover:scale-105 focus:outline-none focus:ring-4 focus:ring-green-400"
      >
        {isLoading ? <img src={spinner} alt="spinner" className="animate-spin w-5 h-5" /> : "Run"}
      </button>
    </div>
  );
};

export default RunInsert;
