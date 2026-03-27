
import { useEffect } from "react";
import { replace, useNavigate } from "react-router-dom";

function Dashboard() {

    const navigate = useNavigate();
    useEffect(() =>{
        const usuario = localStorage.getItem("usuario");

        if(!usuario){
            navigate("/", {replace: true});
        };
    }, [navigate]);

    const logout =() =>{
        localStorage.removeItem("usuario");

        navigate("/",{replace: true})
    };

    return (
        <div>
            <h1>Dashboard</h1>
            <p>Bienvenido al sistema de informes UPS</p>
            <button onClick={logout}>Cerrar Sesión</button>
        </div>
    );
};

export default Dashboard;