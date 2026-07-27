import { useNavigate } from "react-router-dom";
import "../styles/SelectionPage.css";

function ColsofPage() {
    const navigate = useNavigate();
    // Cerrar Sesión
    const logout =() =>{
        localStorage.removeItem("usuario");
        navigate("/",{replace: true})
    };
    const irASelection = () => {
        navigate("/selection");
    }

    return (
        <div className="selection-page">
            <div class="contenedor-boton">
                <button class="boton-cierre-sesion" onClick={irASelection}>Volver a Proyectos</button>
                <button class="boton-cierre-sesion" onClick={logout}>Cerrar Sesión</button>
            </div>
            <h1>Página Colsof</h1>
            <p>Pagina en construcción...</p>
        </div>
    );
}

export default ColsofPage;
