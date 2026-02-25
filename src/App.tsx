import { BrowserRouter, Routes, Route } from "react-router-dom";
import Navbar from "./components/layout/Navbar";
import ProductionPlanMaker from "./components/ProductionPlanMaker";

export default function App() {
  return (
    <BrowserRouter>
      <div className="min-h-screen bg-[#f5eedb] font-manrope">
        <Navbar />
        <main className="pt-20">
          <Routes>
            <Route path="/" element={<ProductionPlanMaker />} />
            <Route path="/production-plan" element={<ProductionPlanMaker />} />
          </Routes>
        </main>
      </div>
    </BrowserRouter>
  );
}
