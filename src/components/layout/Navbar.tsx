import { useState } from "react";
import { Link, useLocation } from "react-router-dom";
import { Menu, X, Home, FileSpreadsheet, ChevronLeft, ChevronRight, History } from "lucide-react";
import { navigation } from "../../data/Navigation";
import logo from "../../assets/lifewood-logo.png";
import icon from "../../assets/icon.png";

// Extend navigation with icons if possible
const navWithIcons = [
  { label: "Home", href: "/", icon: <Home className="w-5 h-5" /> },
  { label: "Production Plan", href: "/production-plan", icon: <FileSpreadsheet className="w-5 h-5" /> },
];

const Navbar = () => {
  const [isExpanded, setIsExpanded] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const location = useLocation();

  return (
    <>
      {/* Desktop Sidebar */}
      <nav
        className={`hidden md:flex flex-col fixed left-0 top-0 h-full bg-[#133020]/90 backdrop-blur-xl border-r border-white/10 shadow-sm z-50 transition-all duration-300 ease-in-out ${isExpanded ? "w-64" : "w-20"}`}
      >
        <div className={`flex items-center h-20 border-b border-white/10 ${isExpanded ? "justify-start px-6 gap-3 bg-white/5" : "justify-center w-full"}`}>
          <img
            src={icon}
            alt="Toggle Sidebar"
            className={`w-10 h-10 object-contain cursor-pointer transition-transform hover:scale-105 shrink-0 ${isExpanded ? "hidden" : "block"}`}
            onClick={() => setIsExpanded(true)}
          />
          <Link
            to="/"
            className={`overflow-hidden transition-all duration-300 flex items-center h-12 cursor-pointer ${isExpanded ? "w-auto opacity-100" : "w-0 opacity-0 hidden"}`}
            onClick={(e) => {
              if (location.pathname === "/") {
                e.preventDefault();
              }
              setIsExpanded(false);
            }}
          >
            <img
              src={logo}
              alt="Lifewood Navigation"
              className="h-8 w-auto object-contain brightness-0 invert min-w-[100px] hover:opacity-80 transition-opacity"
            />
          </Link>
        </div>

        <div className="flex-1 py-6 flex flex-col gap-2 px-3 overflow-y-auto">
          {navWithIcons.map((item) => {
            const isActive = location.pathname === item.href;
            return (
              <Link
                key={item.label}
                to={item.href}
                className={`flex items-center gap-4 px-3 py-3 rounded-xl transition-colors ${isActive
                  ? "bg-[#046241] text-white font-semibold"
                  : "text-white/70 hover:bg-white/10 hover:text-white"
                  }`}
                title={!isExpanded ? item.label : undefined}
              >
                <div className="flex-shrink-0">{item.icon}</div>
                <span className={`whitespace-nowrap transition-opacity duration-300 ${isExpanded ? "opacity-100" : "opacity-0 w-0 hidden"}`}>
                  {item.label}
                </span>
              </Link>
            );
          })}
        </div>

        {/* Toggle History Sidebar */}
        <div className="p-3 border-t border-white/10 mt-auto">
          <button
            onClick={() => window.dispatchEvent(new CustomEvent('toggle-chat-history'))}
            className={`flex items-center gap-4 px-3 py-3 w-full rounded-xl transition-colors text-white/70 hover:bg-white/10 hover:text-white`}
            title={!isExpanded ? "Chat History" : undefined}
          >
            <div className="flex-shrink-0"><History className="w-5 h-5" /></div>
            <span className={`whitespace-nowrap transition-opacity duration-300 ${isExpanded ? "opacity-100" : "opacity-0 w-0 hidden"}`}>
              Chat History
            </span>
          </button>
        </div>
      </nav>

      {/* Mobile Topbar & Hamburger */}
      <nav className="md:hidden fixed top-0 left-0 right-0 h-16 bg-[#133020]/90 backdrop-blur-xl border-b border-white/10 z-50 flex items-center justify-between px-4">
        <Link to="/" className="flex items-center shrink-0">
          <img
            src={logo}
            alt="Lifewood"
            className="h-8 w-auto object-contain brightness-0 invert"
          />
        </Link>
        <button
          onClick={() => setMobileOpen(true)}
          className="p-2 rounded-xl hover:bg-white/10 transition-colors text-white"
          aria-label="Open menu"
        >
          <Menu size={24} />
        </button>
      </nav>

      {/* Mobile Menu Overlay */}
      {mobileOpen && (
        <div className="fixed inset-0 z-[100] md:hidden">
          <div
            className="absolute inset-0 bg-[#133020]/40 backdrop-blur-sm"
            onClick={() => setMobileOpen(false)}
          />
          <div className="absolute right-0 top-0 bottom-0 w-[280px] bg-[#133020] shadow-2xl flex flex-col border-l border-white/10">
            <div className="flex items-center justify-between p-4 border-b border-white/10">
              <img
                src={logo}
                alt="Lifewood"
                className="h-8 w-auto object-contain brightness-0 invert"
              />
              <button
                onClick={() => setMobileOpen(false)}
                className="p-2 rounded-xl hover:bg-white/10 transition-colors text-white"
                aria-label="Close menu"
              >
                <X size={20} />
              </button>
            </div>
            <div className="flex flex-col gap-2 p-4">
              {navWithIcons.map((item) => {
                const isActive = location.pathname === item.href;
                return (
                  <Link
                    key={item.label}
                    to={item.href}
                    onClick={() => setMobileOpen(false)}
                    className={`flex items-center gap-3 px-4 py-3 rounded-xl text-sm font-medium transition-colors ${isActive
                      ? "bg-[#046241] text-white font-semibold"
                      : "text-white/70 hover:bg-white/10 hover:text-white"
                      }`}
                  >
                    {item.icon}
                    {item.label}
                  </Link>
                );
              })}
            </div>
          </div>
        </div>
      )}
    </>
  );
};

export default Navbar;
