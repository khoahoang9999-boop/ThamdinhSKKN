import React, { useState } from 'react';
import { Eye, EyeOff, Loader2 } from 'lucide-react';
import { auth } from '../lib/firebase';
import { GoogleAuthProvider, signInWithPopup, signInWithEmailAndPassword, createUserWithEmailAndPassword, sendPasswordResetEmail, signOut } from 'firebase/auth';

interface LoginScreenProps {
  onLogin: () => void;
}

export default function LoginScreen({ onLogin }: LoginScreenProps) {
  const [username, setUsername] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [error, setError] = useState('');
  const [message, setMessage] = useState(() => {
    const msg = sessionStorage.getItem('loginMessage');
    if (msg) {
      sessionStorage.removeItem('loginMessage');
      return msg;
    }
    return '';
  });
  const [rememberMe, setRememberMe] = useState(false);
  const [isRegistering, setIsRegistering] = useState(false);
  const [loading, setLoading] = useState(false);

  const handleLogin = async (e: React.FormEvent) => {
    e.preventDefault();
    setError('');
    setMessage('');
    setLoading(true);
    
    let loginEmail = username.trim();

    if (!loginEmail.includes('@')) {
      loginEmail = `${loginEmail}@admin.com`;
    }

    try {
      if (isRegistering) {
        sessionStorage.setItem('loginMessage', 'Đăng ký thành công! Vui lòng đăng nhập.');
        await createUserWithEmailAndPassword(auth, loginEmail, password);
        await signOut(auth);
        setMessage('Đăng ký thành công! Vui lòng đăng nhập.');
        setIsRegistering(false);
        setPassword('');
      } else {
        await signInWithEmailAndPassword(auth, loginEmail, password);
        onLogin();
      }
    } catch (err: any) {
      sessionStorage.removeItem('loginMessage');
      if (err.code === 'auth/invalid-credential' || err.code === 'auth/user-not-found' || err.code === 'auth/wrong-password' || err.code === 'auth/invalid-login-credentials') {
        setError('Tài khoản hoặc mật khẩu không chính xác');
      } else if (err.code === 'auth/email-already-in-use') {
        setError('Tài khoản này đã được đăng ký');
      } else if (err.code === 'auth/weak-password') {
        setError('Mật khẩu quá yếu, cần ít nhất 6 ký tự');
      } else if (err.code === 'auth/invalid-email') {
        setError('Tài khoản hoặc email không hợp lệ');
      } else if (err.code === 'auth/too-many-requests') {
        setError('Quá nhiều yêu cầu đăng nhập, vui lòng thử lại sau');
      } else if (err.code === 'auth/user-disabled') {
        setError('Tài khoản này đã bị khóa');
      } else if (err.code === 'auth/operation-not-allowed') {
        setError('Phương thức đăng nhập này chưa được kích hoạt');
      } else if (err.code === 'auth/network-request-failed') {
        setError('Lỗi kết nối mạng, vui lòng kiểm tra lại');
      } else {
        setError('Lỗi không xác định: ' + (err.code || 'Vui lòng thử lại'));
      }
    } finally {
      setLoading(false);
    }
  };

  const handleGoogleLogin = async () => {
    setError('');
    setMessage('');
    const provider = new GoogleAuthProvider();
    try {
      await signInWithPopup(auth, provider);
      onLogin();
    } catch (err: any) {
      if (err.code !== 'auth/popup-closed-by-user') {
        setError('Lỗi đăng nhập Google, vui lòng thử lại');
      }
    }
  };

  const handleForgotPassword = async (e: React.MouseEvent) => {
    e.preventDefault();
    if (!username) {
      setError('Vui lòng nhập tài khoản hoặc email để lấy lại mật khẩu');
      return;
    }
    
    setLoading(true);
    setError('');
    setMessage('');
    
    let loginEmail = username.trim();
    if (!loginEmail.includes('@')) {
      loginEmail = `${loginEmail}@admin.com`;
    }

    try {
      await sendPasswordResetEmail(auth, loginEmail);
      setMessage('Email khôi phục mật khẩu đã được gửi. Vui lòng kiểm tra hộp thư của bạn.');
    } catch (err: any) {
      if (err.code === 'auth/user-not-found') {
        setError('Tài khoản/Email này chưa được đăng ký');
      } else if (err.code === 'auth/invalid-email') {
        setError('Tài khoản hoặc email không hợp lệ');
      } else if (err.code === 'auth/too-many-requests') {
        setError('Quá nhiều yêu cầu, vui lòng thử lại sau');
      } else if (err.code === 'auth/network-request-failed') {
        setError('Lỗi kết nối mạng, vui lòng kiểm tra lại');
      } else {
        setError('Lỗi không xác định: ' + (err.code || 'Vui lòng thử lại'));
      }
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-md p-8">
        <h1 className="text-2xl font-bold text-center text-slate-900 mb-6">
          {isRegistering ? 'Đăng Ký Tài Khoản' : 'Đăng Nhập'}
        </h1>
        
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
              required
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
                required
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

          {message && (
            <p className="text-sm text-emerald-600 font-medium bg-emerald-50 p-3 rounded-lg border border-emerald-100">{message}</p>
          )}

          {!isRegistering && (
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
          )}

          <button
            type="submit"
            disabled={loading}
            className="w-full py-2.5 px-4 bg-blue-600 hover:bg-blue-700 text-white font-bold rounded-lg transition-colors shadow-sm flex justify-center items-center gap-2 disabled:opacity-70"
          >
            {loading && <Loader2 className="w-4 h-4 animate-spin" />}
            {isRegistering ? 'Đăng Ký' : 'Đăng Nhập'}
          </button>

          {!isRegistering && (
            <div className="text-center mt-2">
              <a href="#" onClick={handleForgotPassword} className="text-sm text-blue-600 font-medium hover:underline">
                Quên mật khẩu?
              </a>
            </div>
          )}

          <div className="relative flex items-center py-1">
            <div className="flex-grow border-t border-slate-200"></div>
            <span className="flex-shrink-0 mx-4 text-slate-400 text-sm">Hoặc</span>
            <div className="flex-grow border-t border-slate-200"></div>
          </div>

          <button
            type="button"
            onClick={handleGoogleLogin}
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
            <span className="text-slate-500 text-sm">
              {isRegistering ? 'Đã có tài khoản? ' : 'Chưa có tài khoản? '}
            </span>
            <button
              type="button"
              onClick={() => {
                setIsRegistering(!isRegistering);
                setError('');
              }}
              className="text-blue-600 text-sm font-bold hover:underline"
            >
              {isRegistering ? 'Đăng nhập ngay' : 'Đăng ký ngay'}
            </button>
          </div>
        </form>
      </div>
    </div>
  );
}
