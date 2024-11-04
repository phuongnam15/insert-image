import { createContext, useContext, useEffect, useState } from "react";

const notifyContext = createContext();

const Notification = ({ open, handleClose, msg, color }) => {
  useEffect(() => {
    setTimeout(() => handleClose(), 3000);
  }, []);
  return (
    <>
      <div
        className={`fixed right-2 top-6 z-40 space-y-2 font-popi transition-all duration-500 ${open ? "" : "hidden"}`}
      >
        <div
          className="max-w-xs rounded-xl border border-gray-200 bg-gray-100 shadow-lg"
          role="alert"
          tabIndex="-1"
          aria-labelledby="hs-toast-success-example-label"
        >
          <div className="flex p-4">
            <div className="shrink-0">
              {color == "success" ? (
                <svg
                  className="mt-0.5 size-4 shrink-0 text-teal-500"
                  xmlns="http://www.w3.org/2000/svg"
                  width="16"
                  height="16"
                  fill="currentColor"
                  viewBox="0 0 16 16"
                >
                  <path d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zm-3.97-3.03a.75.75 0 0 0-1.08.022L7.477 9.417 5.384 7.323a.75.75 0 0 0-1.06 1.06L6.97 11.03a.75.75 0 0 0 1.079-.02l3.992-4.99a.75.75 0 0 0-.01-1.05z"></path>
                </svg>
              ) : color == "error" ? (
                <svg
                  className="mt-0.5 size-4 shrink-0 text-red-500"
                  xmlns="http://www.w3.org/2000/svg"
                  width="16"
                  height="16"
                  fill="currentColor"
                  viewBox="0 0 16 16"
                >
                  <path d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zM5.354 4.646a.5.5 0 1 0-.708.708L7.293 8l-2.647 2.646a.5.5 0 0 0 .708.708L8 8.707l2.646 2.647a.5.5 0 0 0 .708-.708L8.707 8l2.647-2.646a.5.5 0 0 0-.708-.708L8 7.293 5.354 4.646z"></path>
                </svg>
              ) : (
                <svg
                  className="mt-0.5 size-4 shrink-0 text-yellow-500"
                  xmlns="http://www.w3.org/2000/svg"
                  width="16"
                  height="16"
                  fill="currentColor"
                  viewBox="0 0 16 16"
                >
                  <path d="M16 8A8 8 0 1 1 0 8a8 8 0 0 1 16 0zM8 4a.905.905 0 0 0-.9.995l.35 3.507a.552.552 0 0 0 1.1 0l.35-3.507A.905.905 0 0 0 8 4zm.002 6a1 1 0 1 0 0 2 1 1 0 0 0 0-2z"></path>
                </svg>
              )}
            </div>
            <div className="ms-3">
              <p className="mr-4 text-sm text-gray-700">
                {msg}
              </p>
            </div>
            <button
              type="button"
              className="-mx-1.5 -my-1.5 ms-auto inline-flex h-8 w-8 items-center justify-center rounded-lg bg-white p-1.5 text-gray-400 hover:bg-gray-100 hover:text-gray-900 focus:ring-2 focus:ring-gray-300"
              data-dismiss-target="#toast-success"
              aria-label="Close"
              onClick={handleClose}
            >
              <svg
                className="h-2 w-2"
                aria-hidden="true"
                xmlns="http://www.w3.org/2000/svg"
                fill="none"
                viewBox="0 0 14 14"
              >
                <path
                  stroke="currentColor"
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth="2"
                  d="m1 1 6 6m0 0 6 6M7 7l6-6M7 7l-6 6"
                />
              </svg>
            </button>
          </div>
        </div>
      </div>
    </>
  );
};

export const NotifyContextProvider = ({ children }) => {
  const [loading, setLoading] = useState(false);
  const [color, setColor] = useState("success");
  const [msg, setMsg] = useState("");

  const show = (msg, color) => {
    setMsg(msg);
    setColor(color);
    setLoading(true);
  };
  const hide = () => {
    setLoading(false);
  };

  return (
    <notifyContext.Provider
      value={{
        loading: loading,
        show: show,
        hide: hide,
      }}
    >
      {loading && (
        <Notification
          open={loading}
          handleClose={hide}
          msg={msg}
          color={color}
        />
      )}
      {children}
    </notifyContext.Provider>
  );
};

export const useNotifyContext = () => useContext(notifyContext);
