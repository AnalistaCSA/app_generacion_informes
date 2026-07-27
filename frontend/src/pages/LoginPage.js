import { useNavigate } from "react-router-dom";
import Login from "../components/Login";
import "../styles/LoginPages.css";

function LoginPage() {
    const navigate = useNavigate();

    const handleLogin = (user) => {
        localStorage.setItem("usuario", JSON.stringify(user));
        navigate("/selection", { replace: true });
    };

    return <Login onLogin={handleLogin} />;
    }

export default LoginPage;