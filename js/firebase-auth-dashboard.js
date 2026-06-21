
    import { initializeApp } from "https://www.gstatic.com/firebasejs/12.10.0/firebase-app.js";
    import { getAuth, onAuthStateChanged, signInWithEmailAndPassword, createUserWithEmailAndPassword, updateProfile, signOut, sendEmailVerification, sendPasswordResetEmail, GoogleAuthProvider, signInWithPopup, EmailAuthProvider, linkWithCredential } from "https://www.gstatic.com/firebasejs/12.10.0/firebase-auth.js";
    import { getFirestore, doc, setDoc, getDoc, deleteDoc, serverTimestamp } from "https://www.gstatic.com/firebasejs/12.10.0/firebase-firestore.js";
    import { getAnalytics } from "https://www.gstatic.com/firebasejs/12.10.0/firebase-analytics.js";

    const firebaseConfig = {
      apiKey: "AIzaSyBtkjBDGp-piXXdpiVgKEzvPmGZNzd6LMI",
      authDomain: "aoyama-nogizaka-activist.firebaseapp.com",
      projectId: "aoyama-nogizaka-activist",
      storageBucket: "aoyama-nogizaka-activist.firebasestorage.app",
      messagingSenderId: "384668278205",
      appId: "1:384668278205:web:78b25a473c6305f264c6d4",
      measurementId: "G-Y1JYXRYHV9"
    };

    const app = initializeApp(firebaseConfig);
    const auth = getAuth(app);
    const db = getFirestore(app);
    const analytics = getAnalytics(app);

    // グローバルに公開
    window.firebaseAuth = auth;
    window.firebaseDb = db;

    // 認証状態の監視
    onAuthStateChanged(auth, async (user) => {
      window.currentUser = user;
      if (user) {
        const snap = await getDoc(doc(db, "users", user.uid));
        if (!snap.exists()) {
          await setDoc(doc(db, "users", user.uid), {
            name: user.displayName || "",
            company: "",
            email: user.email || "",
            affiliation: "",
            affiliationCode: "",
            jobTitle: "",
            plan: "free",
            createdAt: serverTimestamp(),
            registeredFrom: location.pathname,
            referrer: document.referrer || "direct"
          });
        }
      }
      if (typeof updateAuthUI === 'function') {
        updateAuthUI(user);
      } else {
        document.addEventListener('DOMContentLoaded', () => {
          if (typeof updateAuthUI === 'function') updateAuthUI(user);
        });
      }
    });

    // ログイン
    window.firebaseLogin = async (email, password) => {
      const cred = await signInWithEmailAndPassword(auth, email, password);
      return cred.user;
    };

    // 新規登録
    window.firebaseRegister = async (email, password, name, company, affiliation, affiliationCode, jobTitle) => {
      const cred = await createUserWithEmailAndPassword(auth, email, password);
      await updateProfile(cred.user, { displayName: name });
      await setDoc(doc(db, "users", cred.user.uid), {
        name, company: company || "", email,
        affiliation: affiliation || "",
        affiliationCode: affiliationCode || "",
        jobTitle: jobTitle || "",
        plan: "free",
        createdAt: serverTimestamp(),
        registeredFrom: location.pathname,
        referrer: document.referrer || "direct"
      });
      await sendEmailVerification(cred.user);
      window.sendWelcomeEmail(email, name);
      return cred.user;
    };

    // パスワードリセット
    window.firebaseResetPassword = (email) => sendPasswordResetEmail(auth, email);

    // ユーザープロフィール取得
    window.firebaseGetProfile = async (uid) => {
      const snap = await getDoc(doc(db, "users", uid));
      return snap.exists() ? snap.data() : null;
    };

    // Google ログイン
    const googleProvider = new GoogleAuthProvider();
    window.firebaseGoogleLogin = async () => {
      const cred = await signInWithPopup(auth, googleProvider);
      const user = cred.user;
      // Firestore にユーザー情報がなければ作成
      const snap = await getDoc(doc(db, "users", user.uid));
      if (!snap.exists()) {
        await setDoc(doc(db, "users", user.uid), {
          name: user.displayName || "",
          company: "",
          email: user.email || "",
          affiliation: "",
          affiliationCode: "",
          jobTitle: "",
          plan: "free",
          createdAt: serverTimestamp(),
          registeredFrom: location.pathname,
          referrer: document.referrer || "direct"
        });
        window.sendWelcomeEmail(user.email, user.displayName || '');
        return { user, needsProfile: true };
      }
      const data = snap.data();
      const needsProfile = !data.affiliation || !data.jobTitle || !data.company;
      return { user, needsProfile };
    };

    // ウェルカムメール送信
    window.sendWelcomeEmail = async (email, name) => {
      try {
        await fetch('https://script.google.com/macros/s/AKfycbyn3g-fn4edjI84f-eJ2UHeuCLlZNbvR4sTuup9qYvwGqBypSh6ndtkYjTR6yBGq4zY/exec', {
          method: 'POST',
          mode: 'no-cors',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ email, name })
        });
      } catch (e) { console.warn('Welcome email failed:', e); }
    };

    // プロフィール更新
    window.firebaseUpdateProfile = async (uid, data) => {
      const ref = doc(db, "users", uid);
      await setDoc(ref, data, { merge: true });
    };

    // パスワードをリンク（Google登録ユーザー向け）
    window.firebaseLinkPassword = async (email, password) => {
      const credential = EmailAuthProvider.credential(email, password);
      await linkWithCredential(auth.currentUser, credential);
    };

    window.firebaseDeleteAccount = async () => {
      const user = auth.currentUser;
      if (!user) throw new Error('ログインされていません');
      await deleteDoc(doc(db, "users", user.uid));
      await user.delete();
    };

    // ログアウト
    window.firebaseLogout = () => signOut(auth);
  