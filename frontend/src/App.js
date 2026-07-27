import { BrowserRouter, Routes, Route } from "react-router-dom";
import LoginPage from "./pages/LoginPage";
import Dashboard from "./pages/Dashboard";
import SelectionPage from "./pages/SelectionPage";
import ColsofPage from "./pages/ColsofPage";

function App() {
  return (
    <BrowserRouter>
      <Routes>
        <Route path="/" element={<LoginPage />} />
        <Route path="/dashboard" element={<Dashboard />} />
        <Route path="/selection" element={<SelectionPage />} />
        <Route path="/colsof" element={<ColsofPage />} />
      </Routes>
    </BrowserRouter>
  );
}

export default App;
