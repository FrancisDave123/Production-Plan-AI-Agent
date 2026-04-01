import { useEffect, useState } from "react";
import { Link, useLocation, useNavigate } from "react-router-dom";
import { Menu, X, Home, FileSpreadsheet, ChevronLeft, ChevronRight, History, User, LogOut, AlertTriangle } from "lucide-react";
import { useAuth } from "../../contexts/AuthContext";
import { motion, AnimatePresence } from "motion/react";
import { navigation } from "../../data/Navigation";
const logo = "https://framerusercontent.com/images/BZSiFYgRc4wDUAuEybhJbZsIBQY.png?width=1519&height=429";
import icon from "../../assets/icon.png";

// Extend navigation with icons if possible
const navWithIcons = [
  { label: "Home", href: "/", icon: <Home className="w-5 h-5" /> },
  { label: "Production Plan", href: "/production-plan", icon: <FileSpreadsheet className="w-5 h-5" /> },
];

const Navbar = () => {
  const [isExpanded, setIsExpanded] = useState(false);
  const [mobileOpen, setMobileOpen] = useState(false);
  const [showSignOutConfirm, setShowSignOutConfirm] = useState(false);
  const { isSignedIn, login, logout } = useAuth();
  const location = useLocation();
  const navigate = useNavigate();

  useEffect(() => {
    const event = new CustomEvent("sidebar-expanded-change", {
      detail: isExpanded,
    });
    window.dispatchEvent(event);
  }, [isExpanded]);

  return (
    <>
      {/* Desktop Sidebar */}
      <nav
        className={`hidden md:flex flex-col fixed left-0 top-0 h-full bg-[#133020]/90 backdrop-blur-xl shadow-sm z-50 overflow-x-hidden transition-all duration-300 ease-in-out ${isExpanded ? "w-64" : "w-16"}`}
      >
        <div className={`flex items-center h-20 border-b border-white/10 ${isExpanded ? "justify-start px-6 gap-3 bg-white/5" : "justify-center w-full"}`}>
          {!isExpanded && (
            <button
              type="button"
              onClick={() => setIsExpanded(true)}
              className="relative inline-flex items-center justify-center w-10 h-10 shrink-0 group transition-transform hover:scale-105"
              aria-label="Expand sidebar"
            >
              <img
                src={icon}
                alt="Toggle Sidebar"
                className="w-10 h-10 object-contain transition-opacity duration-150 group-hover:opacity-0"
              />
              <ChevronRight
                className="absolute w-5 h-5 text-white opacity-0 translate-x-0.5 group-hover:opacity-100 transition-opacity duration-150"
              />
            </button>
          )}
          <Link
            to="/"
            className={`overflow-hidden transition-all duration-300 flex items-center h-12 cursor-pointer ${isExpanded ? "w-auto opacity-100" : "w-0 opacity-0 hidden"}`}
            onClick={(e) => {
              if (location.pathname === "/") {
                e.preventDefault();
              }
            }}
          >
            <img
              src={logo}
              alt="Lifewood Navigation"
              className="h-8 w-auto object-contain brightness-0 invert min-w-[100px] hover:opacity-80 transition-opacity"
            />
          </Link>
          {isExpanded && (
            <button
              type="button"
              onClick={() => setIsExpanded(false)}
              className="ml-auto p-2 rounded-full hover:bg-white/10 text-white transition-colors"
              aria-label="Collapse sidebar"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
          )}
        </div>

        <div
          className={`flex-1 py-6 flex flex-col gap-2 overflow-y-auto ${isExpanded ? "px-3" : "px-0"
            }`}
        >
          {navWithIcons.map((item) => {
            const isActive = location.pathname === item.href;
            return (
              <Link
                key={item.label}
                to={item.href}
                className={`flex items-center rounded-full transition-colors ${isExpanded
                  ? "justify-start gap-4 px-3 py-3"
                  : "justify-center w-12 h-12 mx-auto"
                  } ${isActive
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

        {/* Bottom Actions */}
        <div
          className={`border-t border-white/10 mt-auto flex flex-col gap-2 ${isExpanded ? "p-3" : "px-0 py-3"
            }`}
        >
          {/* Toggle History Sidebar */}
          <button
            onClick={() => {
              if (location.pathname !== "/") {
                navigate("/");
                setTimeout(
                  () =>
                    window.dispatchEvent(
                      new CustomEvent("toggle-chat-history"),
                    ),
                  150,
                );
              } else {
                window.dispatchEvent(
                  new CustomEvent("toggle-chat-history"),
                );
              }
            }}
            className={`flex items-center rounded-full transition-colors text-white/70 hover:bg-white/10 hover:text-white ${isExpanded
              ? "justify-start gap-4 px-3 py-3 w-full"
              : "justify-center w-12 h-12 mx-auto"
              }`}
            title={!isExpanded ? "Chat History" : undefined}
          >
            <div className="flex-shrink-0"><History className="w-5 h-5" /></div>
            <span className={`whitespace-nowrap transition-opacity duration-300 ${isExpanded ? "opacity-100" : "opacity-0 w-0 hidden"}`}>
              Chat History
            </span>
          </button>

          {/* Account Sign In/Out */}
          <button
            onClick={() => {
              if (isSignedIn) {
                setShowSignOutConfirm(true);
              } else {
                login();
              }
            }}
            className={`flex items-center rounded-full transition-colors text-white/70 hover:bg-white/10 hover:text-white ${isExpanded
              ? "justify-start gap-4 px-3 py-3 w-full"
              : "justify-center w-12 h-12 mx-auto"
              }`}
            title={!isExpanded ? (isSignedIn ? "Sign Out" : "Sign In") : undefined}
          >
            <div className="flex-shrink-0">
              {isSignedIn ? <LogOut className="w-5 h-5" /> : <User className="w-5 h-5" />}
            </div>
            <span className={`whitespace-nowrap transition-opacity duration-300 ${isExpanded ? "opacity-100" : "opacity-0 w-0 hidden"}`}>
              {isSignedIn ? "Sign Out" : "Sign In"}
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
                    className={`flex items-center gap-3 px-4 py-3 rounded-full text-sm font-medium transition-colors ${isActive
                      ? "bg-[#046241] text-white font-semibold"
                      : "text-white/70 hover:bg-white/10 hover:text-white"
                      }`}
                  >
                    {item.icon}
                    {item.label}
                  </Link>
                );
              })}

              <div className="w-full h-px bg-white/10 my-2"></div>

              <button
                onClick={() => {
                  if (isSignedIn) {
                    setShowSignOutConfirm(true);
                  } else {
                    login();
                  }
                  setMobileOpen(false);
                }}
                className="flex items-center gap-3 px-4 py-3 rounded-full text-sm font-medium transition-colors text-white/70 hover:bg-white/10 hover:text-white w-full text-left"
              >
                {isSignedIn ? <LogOut className="w-5 h-5" /> : <User className="w-5 h-5" />}
                {isSignedIn ? "Sign Out" : "Sign In"}
              </button>
            </div>
          </div>
        </div>
      )}
      {/* Sign Out Confirmation Modal */}
      <AnimatePresence>
        {showSignOutConfirm && (
          <div className="fixed inset-0 z-[1000] flex items-center justify-center p-4">
            <motion.div
              initial={{ opacity: 0 }}
              animate={{ opacity: 1 }}
              exit={{ opacity: 0 }}
              onClick={() => setShowSignOutConfirm(false)}
              className="absolute inset-0 bg-black/60 backdrop-blur-sm"
            />
            <motion.div
              initial={{ opacity: 0, scale: 0.95, y: 20 }}
              animate={{ opacity: 1, scale: 1, y: 0 }}
              exit={{ opacity: 0, scale: 0.95, y: 20 }}
              className="relative w-full max-w-sm bg-[#133020] border border-white/10 rounded-2xl shadow-2xl overflow-hidden"
            >
              <div className="p-6">
                <div className="flex items-center justify-center w-12 h-12 mb-4 rounded-full bg-red-500/10 text-red-500">
                  <AlertTriangle className="w-6 h-6" />
                </div>
                <h3 className="text-xl font-semibold text-white mb-2">Sign Out</h3>
                <p className="text-white/60">
                  Are you sure you want to sign out of your account?
                </p>
              </div>
              <div className="flex border-t border-white/10">
                <button
                  onClick={() => setShowSignOutConfirm(false)}
                  className="flex-1 px-4 py-4 text-sm font-medium text-white/70 hover:bg-white/5 transition-colors"
                >
                  Cancel
                </button>
                <button
                  onClick={() => {
                    logout();
                    setShowSignOutConfirm(false);
                  }}
                  className="flex-1 px-4 py-4 text-sm font-medium text-red-400 hover:bg-red-500/10 transition-colors border-l border-white/10"
                >
                  Sign Out
                </button>
              </div>
            </motion.div>
          </div>
        )}
      </AnimatePresence>
    </>
  );
};

export default Navbar;
