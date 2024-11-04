import { useRoutes } from "react-router-dom";
import Router from "./routers/Router";
import { Suspense } from "react";
import Spinner from "./pages/spinner/Spinner";

const App = () => {
  const routes = useRoutes(Router);
  return <Suspense fallback={<Spinner />}>{routes}</Suspense>;
};

export default App;
