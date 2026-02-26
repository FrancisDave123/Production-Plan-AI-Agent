import { useState } from "react";
import { Link, useLocation } from "react-router-dom";
import { Menu, X } from "lucide-react";
import { navigation } from "../../data/Navigation";
import logo from "../../assets/lifewood-logo.png";

const Navbar = () => {
  const [mobileOpen, setMobileOpen] = useState(false);
  const location = useLocation();

  return (
    <>
      <nav className="fixed top-0 left-0 right-0 z-50 flex justify-center px-6 pt-5 pb-2">
        <div className="bg-white/90 backdrop-blur-xl border border-white/40 shadow-sm rounded-full px-6 py-3 flex items-center justify-between w-full max-w-4xl">
          <Link to="/" className="flex items-center shrink-0">
            <img
              src={logo}
              alt="Lifewood"
              className="h-8 w-auto object-contain"
            />
          </Link>

          <div className="hidden md:flex items-center gap-6">
            {navigation.map((item) => {
              const isActive = location.pathname === item.href;
              return (
                <Link
                  key={item.label}
                  to={item.href}
                  className={`text-sm font-medium transition-colors ${isActive
                      ? "text-[#046241] font-semibold"
                      : "text-[#133020] hover:text-[#046241]"
                    }`}
                >
                  {item.label}
                </Link>
              );
            })}
          </div>

          <button
            onClick={() => setMobileOpen(true)}
            className="md:hidden p-2 rounded-full hover:bg-gray-100 transition-colors text-[#133020]"
            aria-label="Open menu"
            data-testid="mobile-menu-button"
          >
            <Menu size={20} />
          </button>
        </div>
      </nav>

      {mobileOpen && (
        <div className="fixed inset-0 z-[100] md:hidden">
          <div
            className="absolute inset-0 bg-[#133020]/20"
            onClick={() => setMobileOpen(false)}
          />
          <div className="absolute bottom-0 left-0 right-0 bg-white rounded-t-3xl shadow-2xl p-6">
            <div className="w-10 h-1 bg-gray-200 rounded-full mx-auto mb-6" />
            <div className="flex items-center justify-between mb-6">
              <img
                src={logo}
                alt="Lifewood"
                className="h-8 w-auto object-contain"
              />
              <button
                onClick={() => setMobileOpen(false)}
                className="p-2 rounded-xl hover:bg-gray-100 transition-colors"
                aria-label="Close menu"
                data-testid="mobile-menu-close"
              >
                <X size={20} className="text-[#133020]" />
              </button>
            </div>
            <div className="flex flex-col gap-2">
              {navigation.map((item) => {
                const isActive = location.pathname === item.href;
                return (
                  <Link
                    key={item.label}
                    to={item.href}
                    onClick={() => setMobileOpen(false)}
                    className={`px-4 py-3 rounded-xl text-sm font-medium transition-colors ${isActive
                        ? "bg-[#046241]/10 text-[#046241] font-semibold"
                        : "text-[#133020] hover:bg-gray-50"
                      }`}
                    data-testid={`mobile-nav-${item.label.toLowerCase().replace(/\s+/g, "-")}`}
                  >
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
