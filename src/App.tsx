import { BrowserRouter, Routes, Route } from "react-router-dom";
import Navbar from "./components/layout/Navbar";
import ProductionPlanMaker from "./components/ProductionPlanMaker";

export default function App() {
  return (
    <BrowserRouter>
      <div className="bg-white font-manrope h-screen flex">
        <Navbar />
        <main className="flex-1 transition-all duration-300 pt-16 md:pt-0 md:pl-20">
          <Routes>
            <Route path="/" element={<ProductionPlanMaker />} />
            <Route path="/production-plan" element={<ProductionPlanMaker />} />
          </Routes>
        </main>
      </div>
    </BrowserRouter>
  );
}
