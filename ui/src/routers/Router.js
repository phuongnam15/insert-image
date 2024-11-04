import { lazy } from "react";

const RunInsert = lazy(() => import("../pages/RunInsert"));

const Router = [
  {
    path: "/",
    element: <RunInsert />,
  },
];

export default Router;
