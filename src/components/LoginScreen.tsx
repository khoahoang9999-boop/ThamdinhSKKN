import React, { useState } from 'react';
import { Eye, EyeOff } from 'lucide-react';

interface LoginScreenProps {
  onLogin: () => void;
}

export default function LoginScreen({ onLogin }: LoginScreenProps) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [rememberMe, setRememberMe] = useState(false);

  const handleLogin = (e: React.FormEvent) => {
    e.preventDefault();
    if (username === 'admin' && password === '123456') {
      setError('');
      onLogin();
    } else {
      setError('Tài khoản hoặc mật khẩu không chính xác');
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-8">
        <h1 className="text-2xl font-bold text-center text-slate-900 mb-6">Đăng Nhập</h1>
        
        <form onSubmit={handleLogin} className="space-y-4">
          <div>
            <label className="block text-sm font-bold text-red-600 mb-1">
              Tài khoản / Email
            </label>
            <input
              type="text"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              placeholder="Nhập tài khoản hoặc email của bạn..."
              className="w-full px-4 py-2 rounded-lg border border-slate-200 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500 transition-colors placeholder:text-slate-400"
            />
          </div>

          <div>
            <label className="block text-sm font-bold text-red-600 mb-1">
              Mật khẩu
            </label>
            <div className="relative">
              <input
                type={showPassword ? 'text' : 'password'}
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                placeholder="Nhập mật khẩu..."
                className="w-full px-4 py-2 rounded-lg border border-blue-400 focus:outline-none focus:ring-2 focus:ring-blue-500/50 focus:border-blue-500 transition-colors pr-12 placeholder:text-slate-400 bg-blue-50/10"
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-500 hover:text-slate-700 transition-colors p-1"
              >
                {showPassword ? <EyeOff size={18} /> : <Eye size={18} />}
              </button>
            </div>
          </div>

          {error && (
            <p className="text-sm text-red-500 font-medium">{error}</p>
          )}

          <div className="flex items-center gap-2">
            <input
              type="checkbox"
              id="remember"
              checked={rememberMe}
              onChange={(e) => setRememberMe(e.target.checked)}
              className="w-4 h-4 rounded text-blue-600 border-slate-300 focus:ring-blue-500 cursor-pointer"
            />
            <label htmlFor="remember" className="text-sm font-bold text-slate-700 cursor-pointer select-none">
              Nhớ tài khoản và mật khẩu
            </label>
          </div>

          <button
            type="submit"
            className="w-full py-2.5 px-4 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg transition-colors shadow-sm"
          >
            Đăng Nhập
          </button>

          <div className="text-center mt-2">
            <a href="#" className="text-sm text-blue-600 font-medium hover:underline">
              Quên mật khẩu?
            </a>
          </div>

          <div className="relative flex items-center py-1">
            <div className="flex-grow border-t border-slate-200"></div>
            <span className="flex-shrink-0 mx-4 text-slate-400 text-sm">Hoặc</span>
            <div className="flex-grow border-t border-slate-200"></div>
          </div>

          <button
            type="button"
            className="w-full flex items-center justify-center gap-3 py-2 px-4 bg-white border border-slate-200 rounded-lg text-slate-700 font-bold hover:bg-slate-50 transition-colors"
          >
            <svg viewBox="0 0 24 24" className="w-5 h-5">
              <path d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z" fill="#4285F4" />
              <path d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z" fill="#34A853" />
              <path d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z" fill="#FBBC05" />
              <path d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z" fill="#EA4335" />
            </svg>
            Google
          </button>

          <div className="text-center">
            <span className="text-slate-500 text-sm">Chưa có tài khoản? </span>
            <a href="#" className="text-blue-600 text-sm font-bold hover:underline">
              Đăng ký ngay
            </a>
          </div>
        </form>
      </div>
    </div>
  );
}
