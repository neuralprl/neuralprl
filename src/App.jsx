import React, { useState, useMemo } from 'react';
import { 
  Building2, 
  FileText, 
  Upload, 
  Users, 
  AlertTriangle, 
  CheckCircle2, 
  XCircle, 
  Search, 
  ExternalLink, 
  LogOut, 
  Plus, 
  FolderOpen, 
  ShieldCheck, 
  FileSpreadsheet, 
  Trash2, 
  Link as LinkIcon,
  HardHat,
  Bell,
  Check,
  Clock,
  Send,
  Eye
} from 'lucide-react';
import * as XLSX from 'xlsx';

// --- DATOS INICIALES ---

// Usuarios
// Roles: 'superadmin' (SA), 'corporativo' (UC), 'externo' (UX)
const INITIAL_USERS = [
  { id: '1', email: 'neuralprl', code: 'Neuralprl@', name: 'Superadministrador', role: 'superadmin', assignedCentres: ['ALL'], company: 'Neural PRL' },
  { id: '2', email: 'director.madrid@neural.es', code: 'Pass1234@', name: 'Carlos (Director Madrid)', role: 'corporativo', assignedCentres: ['c1'], company: 'Neural SRL' },
  { id: '3', email: 'prevencion@contratasvalencia.com', code: 'Externa123@', name: 'Mantenimientos Levante SL', role: 'externo', assignedCentres: ['c2'], company: 'Mantenimientos Levante SL' }
];

const INITIAL_GENERAL_DOCS = [
  { id: 'gd1', title: 'Procedimiento General de Evacuación v2', category: 'Procedimientos', link: 'https://sharepoint.com/doc1', date: '2026-01-15' },
  { id: 'gd2', title: 'Protocolo de Actuación Accidentes Laborales', category: 'Protocolos', link: 'https://sharepoint.com/doc2', date: '2026-02-01' },
  { id: 'gd3', title: 'Plantilla de Inspección de Equipos de Protección', category: 'Plantillas', link: 'https://sharepoint.com/doc3', date: '2026-02-10' },
];

const INITIAL_CENTRES = [
  {
    id: 'c1',
    name: 'Centro Neural Madrid - Castellana',
    zone: 'Madrid Norte',
    users: ['neuralprl', 'director.madrid@neural.es'],
    docs: {
      evaluacion_riesgos: [{ id: 'd1', name: 'Evaluacion_Riesgos_2026_Madrid.pdf', link: 'https://sharepoint.com/eval-madrid.pdf' }],
      informacion_riesgos: [{ id: 'd2', name: 'Info_Riesgos_Puestos_Madrid.pdf', link: 'https://sharepoint.com/info-madrid.pdf' }],
      medidas_emergencia: []
    }
  },
  {
    id: 'c2',
    name: 'Centro Neural Valencia - Mestalla',
    zone: 'Comunidad Valenciana',
    users: ['neuralprl', 'prevencion@contratasvalencia.com'],
    docs: {
      evaluacion_riesgos: [{ id: 'd3', name: 'Evaluacion_Riesgos_Mestalla_v1.pdf', link: 'https://sharepoint.com/eval-valencia.pdf' }],
      informacion_riesgos: [],
      medidas_emergencia: [{ id: 'd4', name: 'Plan_Emergencia_Valencia_2026.pdf', link: 'https://sharepoint.com/emerg-valencia.pdf' }]
    }
  }
];

// Registros de Coordinación CAE
// status: 'pendiente_lectura' | 'documentos_enviados' | 'completado'
const INITIAL_CAE_RECORDS = [
  {
    id: 'cae_1',
    centreId: 'c2',
    companyName: 'Mantenimientos Levante SL',
    userEmail: 'prevencion@contratasvalencia.com',
    docsRead: false,
    docsSent: false,
    status: 'pendiente_lectura',
    updatedAt: '2026-02-20'
  }
];

// Notificaciones del correo simulado para el SA
const INITIAL_NOTIFICATIONS = [];

const extractFileNameFromUrl = (url) => {
  if (!url) return '';
  try {
    const parsed = new URL(url);
    const pathname = parsed.pathname;
    const filename = pathname.split('/').pop();
    if (filename && filename.length > 0) return decodeURIComponent(filename);
  } catch (e) {
    const parts = url.split('/');
    const last = parts.pop() || parts.pop();
    if (last) return decodeURIComponent(last.split('?')[0]);
  }
  return 'Documento SharePoint';
};

export default function App() {
  // Autenticación
  const [currentUser, setCurrentUser] = useState(null);
  const [loginEmail, setLoginEmail] = useState('');
  const [loginCode, setLoginCode] = useState('');
  const [loginError, setLoginError] = useState('');

  // Estado Global
  const [users, setUsers] = useState(INITIAL_USERS);
  const [centres, setCentres] = useState(INITIAL_CENTRES);
  const [generalDocs, setGeneralDocs] = useState(INITIAL_GENERAL_DOCS);
  const [caeRecords, setCaeRecords] = useState(INITIAL_CAE_RECORDS);
  const [notifications, setNotifications] = useState(INITIAL_NOTIFICATIONS);

  // Navegación
  const [activeTab, setActiveTab] = useState('centres');
  const [selectedCentre, setSelectedCentre] = useState(null);
  const [searchTerm, setSearchTerm] = useState('');

  // Modal para SharePoint
  const [editDocModal, setEditDocModal] = useState({ open: false, centreId: null, categoryKey: null, categoryLabel: '' });
  const [newLinkUrl, setNewLinkUrl] = useState('');
  const [newLinkName, setNewLinkName] = useState('');

  // Manejo de Login
  const handleLogin = (e) => {
    e.preventDefault();
    setLoginError('');
    const user = users.find(
      u => u.email.trim().toLowerCase() === loginEmail.trim().toLowerCase() && 
           u.code.trim() === loginCode.trim()
    );

    if (user) {
      setCurrentUser(user);
      // Redirigir a vista según rol por defecto
      if (user.role === 'externo') {
        setActiveTab('cae');
      } else {
        setActiveTab('centres');
      }
    } else {
      setLoginError('Usuario o contraseña incorrectos.');
    }
  };

  const handleLogout = () => {
    setCurrentUser(null);
    setLoginEmail('');
    setLoginCode('');
  };

  // Filtrar centros accesibles según el rol
  const accessibleCentres = useMemo(() => {
    if (!currentUser) return [];
    if (currentUser.role === 'superadmin') return centres;
    return centres.filter(c => c.users.includes(currentUser.email));
  }, [currentUser, centres]);

  const currentSelectedCentre = useMemo(() => {
    if (!selectedCentre) return null;
    return centres.find(c => c.id === selectedCentre.id) || selectedCentre;
  }, [centres, selectedCentre]);

  // Carga Masiva Excel
  const handleFileUploadCentres = (e) => {
    const file = e.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const bstr = evt.target.result;
      const wb = XLSX.read(bstr, { type: 'binary' });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const data = XLSX.utils.sheet_to_json(ws, { header: 1 });

      const rows = data.slice(1);
      const newCentres = [...centres];
      const newUsers = [...users];

      rows.forEach((row, idx) => {
        if (!row[0]) return;

        const centreName = row[0].toString().trim();
        const zone = row[1] ? row[1].toString().trim() : 'General';
        const assignedUserEmails = [];

        const userCols = [
          { emailIdx: 2, codeIdx: 3 },
          { emailIdx: 4, codeIdx: 5 },
          { emailIdx: 6, codeIdx: 7 },
          { emailIdx: 8, codeIdx: 9 },
          { emailIdx: 10, codeIdx: 11 }
        ];

        userCols.forEach(({ emailIdx, codeIdx }) => {
          if (row[emailIdx] && row[codeIdx]) {
            const email = row[emailIdx].toString().trim();
            const code = row[codeIdx].toString().trim();
            assignedUserEmails.push(email);

            const exists = newUsers.some(u => u.email.toLowerCase() === email.toLowerCase());
            if (!exists) {
              newUsers.push({
                id: `u_${Date.now()}_${Math.random()}`,
                email: email,
                code: code,
                name: email.includes('@') ? email.split('@')[0] : email,
                role: 'corporativo',
                assignedCentres: []
              });
            }
          }
        });

        const existingIndex = newCentres.findIndex(c => c.name.toLowerCase() === centreName.toLowerCase());
        const centreObj = {
          id: existingIndex >= 0 ? newCentres[existingIndex].id : `c_${Date.now()}_${idx}`,
          name: centreName,
          zone: zone,
          users: assignedUserEmails,
          docs: existingIndex >= 0 ? newCentres[existingIndex].docs : {
            evaluacion_riesgos: [],
            informacion_riesgos: [],
            medidas_emergencia: []
          }
        };

        if (existingIndex >= 0) {
          newCentres[existingIndex] = centreObj;
        } else {
          newCentres.push(centreObj);
        }
      });

      setCentres(newCentres);
      setUsers(newUsers);
      alert('¡Centros y usuarios importados correctamente!');
    };
    reader.readAsBinaryString(file);
  };

  // --- LÓGICA MÓDULO CAE ---

  // Marcar lectura de documentos (UX)
  const handleToggleReadDocs = (centreId) => {
    setCaeRecords(prev => {
      const existing = prev.find(r => r.centreId === centreId && r.userEmail === currentUser.email);
      if (existing) {
        return prev.map(r => r.id === existing.id ? { ...r, docsRead: !r.docsRead } : r);
      } else {
        return [...prev, {
          id: `cae_${Date.now()}`,
          centreId,
          companyName: currentUser.company || currentUser.name,
          userEmail: currentUser.email,
          docsRead: true,
          docsSent: false,
          status: 'pendiente_lectura',
          updatedAt: new Date().toISOString().split('T')[0]
        }];
      }
    });
  };

  // Enviar documentación de coordinación (UX) -> Genera email/notificación al SA
  const handleSendCaeDocs = (centreId) => {
    const centre = centres.find(c => c.id === centreId);
    
    setCaeRecords(prev => {
      const existing = prev.find(r => r.centreId === centreId && r.userEmail === currentUser.email);
      if (existing) {
        return prev.map(r => r.id === existing.id ? { 
          ...r, 
          docsSent: true, 
          status: 'documentos_enviados',
          updatedAt: new Date().toISOString().split('T')[0]
        } : r);
      } else {
        return [...prev, {
          id: `cae_${Date.now()}`,
          centreId,
          companyName: currentUser.company || currentUser.name,
          userEmail: currentUser.email,
          docsRead: true,
          docsSent: true,
          status: 'documentos_enviados',
          updatedAt: new Date().toISOString().split('T')[0]
        }];
      }
    });

    // Simulación de correo enviado al Superadministrador
    const newNotif = {
      id: `notif_${Date.now()}`,
      title: 'Nueva Documentación CAE Enviada',
      message: `La empresa ${currentUser.company || currentUser.name} (${currentUser.email}) ha enviado la documentación CAE para el centro ${centre?.name || ''}.`,
      date: new Date().toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }),
      read: false
    };
    setNotifications(prev => [newNotif, ...prev]);

    alert('Documentación de coordinación enviada con éxito. Se ha notificado al Superadministrador para su validación.');
  };

  // Validar y Aprobar Acceso (Superadministrador SA)
  const handleApproveCae = (recordId) => {
    setCaeRecords(prev => prev.map(r => {
      if (r.id === recordId) {
        return { ...r, status: 'completado' };
      }
      return r;
    }));
  };

  // Modal Enlaces SharePoint
  const handleUrlChange = (url) => {
    setNewLinkUrl(url);
    if (url.trim() && !newLinkName) {
      setNewLinkName(extractFileNameFromUrl(url));
    }
  };

  const handleAddLink = (e) => {
    e.preventDefault();
    if (!newLinkUrl.trim() || !editDocModal.centreId || !editDocModal.categoryKey) return;

    const fileName = newLinkName.trim() || extractFileNameFromUrl(newLinkUrl);
    const newDocObj = {
      id: `doc_${Date.now()}`,
      name: fileName,
      link: newLinkUrl.trim()
    };

    setCentres(prevCentres => prevCentres.map(c => {
      if (c.id === editDocModal.centreId) {
        const currentList = Array.isArray(c.docs[editDocModal.categoryKey]) 
          ? c.docs[editDocModal.categoryKey] 
          : [];
        return {
          ...c,
          docs: {
            ...c.docs,
            [editDocModal.categoryKey]: [...currentList, newDocObj]
          }
        };
      }
      return c;
    }));

    setNewLinkUrl('');
    setNewLinkName('');
  };

  const handleDeleteLink = (docId) => {
    setCentres(prevCentres => prevCentres.map(c => {
      if (c.id === editDocModal.centreId) {
        const currentList = Array.isArray(c.docs[editDocModal.categoryKey]) 
          ? c.docs[editDocModal.categoryKey] 
          : [];
        return {
          ...c,
          docs: {
            ...c.docs,
            [editDocModal.categoryKey]: currentList.filter(d => d.id !== docId)
          }
        };
      }
      return c;
    }));
  };

  // Pantalla de Login
  if (!currentUser) {
    return (
      <div className="min-h-screen bg-slate-900 flex items-center justify-center p-4 font-sans">
        <div className="bg-white rounded-xl shadow-2xl p-8 max-w-md w-full space-y-6">
          <div className="text-center space-y-2">
            <div className="bg-blue-600 text-white w-12 h-12 rounded-lg flex items-center justify-center mx-auto shadow-lg">
              <ShieldCheck className="w-8 h-8" />
            </div>
            <h1 className="text-2xl font-bold text-slate-800">Plataforma PRL & CAE</h1>
            <p className="text-sm text-slate-500">Gestión Documental y Coordinación Empresarial</p>
          </div>

          <form onSubmit={handleLogin} className="space-y-4">
            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Usuario / Correo</label>
              <input 
                type="text" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm"
                placeholder="neuralprl o correo@empresa.com"
                value={loginEmail}
                onChange={(e) => setLoginEmail(e.target.value)}
              />
            </div>

            <div>
              <label className="block text-xs font-semibold text-slate-600 uppercase mb-1">Contraseña</label>
              <input 
                type="password" 
                required
                className="w-full px-4 py-2 border border-slate-300 rounded-lg focus:ring-2 focus:ring-blue-500 focus:outline-none text-sm"
                placeholder="Tu contraseña"
                value={loginCode}
                onChange={(e) => setLoginCode(e.target.value)}
              />
            </div>

            {loginError && (
              <div className="p-3 bg-red-50 border border-red-200 text-red-600 text-xs rounded-lg flex items-center gap-2">
                <AlertTriangle className="w-4 h-4 shrink-0" />
                <span>{loginError}</span>
              </div>
            )}

            <button 
              type="submit" 
              className="w-full py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition shadow-md text-sm"
            >
              Iniciar Sesión
            </button>
          </form>

          <div className="text-xs text-center text-slate-400 border-t pt-4">
            Acceso según rol: Superadministrador (SA), Corporativo (UC) o Empresa Externa (UX).
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans">
      {/* Header Superior */}
      <header className="bg-slate-800 text-white shadow-md sticky top-0 z-40">
        <div className="max-w-7xl mx-auto px-4 py-3 flex justify-between items-center">
          <div className="flex items-center space-x-3">
            <ShieldCheck className="w-7 h-7 text-blue-400" />
            <div>
              <h1 className="font-bold text-lg leading-tight">Neural PRL</h1>
              <span className="text-xs text-slate-400">Coordinación de Actividades Empresariales</span>
            </div>
          </div>

          <div className="flex items-center space-x-4">
            {/* Notificaciones Correo para el Superadministrador */}
            {currentUser.role === 'superadmin' && (
              <div className="relative group">
                <button className="p-2 text-slate-300 hover:text-white rounded-lg relative">
                  <Bell className="w-5 h-5" />
                  {notifications.filter(n => !n.read).length > 0 && (
                    <span className="absolute top-1 right-1 bg-amber-500 text-slate-900 font-bold text-[10px] w-4 h-4 rounded-full flex items-center justify-center">
                      {notifications.filter(n => !n.read).length}
                    </span>
                  )}
                </button>

                {/* Dropdown Notificaciones / Mail SA */}
                <div className="absolute right-0 mt-2 w-80 bg-white rounded-xl shadow-xl border border-slate-200 text-slate-800 p-4 hidden group-hover:block z-50">
                  <h4 className="text-xs font-bold uppercase text-slate-500 mb-2 border-b pb-1 flex justify-between">
                    <span>Buzón del Superadministrador</span>
                    <span className="text-[10px] text-blue-600">Mails CAE</span>
                  </h4>
                  <div className="space-y-2 max-h-60 overflow-y-auto">
                    {notifications.length === 0 ? (
                      <p className="text-xs text-slate-400 italic py-2 text-center">Sin notificaciones nuevas</p>
                    ) : (
                      notifications.map(n => (
                        <div key={n.id} className="p-2 bg-slate-50 rounded border border-slate-100 text-xs">
                          <p className="font-bold text-slate-700">{n.title}</p>
                          <p className="text-slate-500 text-[11px] mt-0.5">{n.message}</p>
                          <span className="text-[9px] text-slate-400 block mt-1">{n.date}</span>
                        </div>
                      ))
                    )}
                  </div>
                </div>
              </div>
            )}

            <div className="text-right hidden sm:block">
              <p className="text-sm font-medium">{currentUser.name}</p>
              <div className="flex items-center justify-end gap-1">
                <span className={`text-[10px] px-2 py-0.5 rounded font-semibold uppercase ${
                  currentUser.role === 'superadmin' ? 'bg-purple-900 text-purple-200' :
                  currentUser.role === 'corporativo' ? 'bg-blue-900 text-blue-200' : 'bg-emerald-900 text-emerald-200'
                }`}>
                  {currentUser.role === 'superadmin' ? 'Superadmin (SA)' :
                   currentUser.role === 'corporativo' ? 'Usuario Corporativo (UC)' : 'Empresa Externa (UX)'}
                </span>
              </div>
            </div>

            <button 
              onClick={handleLogout} 
              className="p-2 text-slate-300 hover:text-white hover:bg-slate-700 rounded-lg transition"
              title="Cerrar Sesión"
            >
              <LogOut className="w-5 h-5" />
            </button>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 py-6 flex-1 w-full flex flex-col md:flex-row gap-6">
        
        {/* NAVEGACIÓN LATERAL / MÓDULO CAE */}
        <aside className="w-full md:w-64 bg-white rounded-xl border border-slate-200 shadow-sm p-4 shrink-0 h-fit space-y-6">
          <div className="space-y-1">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-2">Navegación</span>
            
            <button 
              onClick={() => { setActiveTab('centres'); setSelectedCentre(null); }}
              className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2.5 transition ${activeTab === 'centres' ? 'bg-blue-50 text-blue-600 font-semibold' : 'text-slate-600 hover:bg-slate-50'}`}
            >
              <Building2 className="w-4 h-4" />
              <span>Centros de Trabajo</span>
            </button>

            <button 
              onClick={() => setActiveTab('general')}
              className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2.5 transition ${activeTab === 'general' ? 'bg-blue-50 text-blue-600 font-semibold' : 'text-slate-600 hover:bg-slate-50'}`}
            >
              <FolderOpen className="w-4 h-4" />
              <span>Doc. General PRL</span>
            </button>
          </div>

          {/* MÓDULO DESTACADO LATERAL: CAE */}
          <div className="pt-4 border-t border-slate-100 space-y-2">
            <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-2">Módulo Especial</span>
            
            <button 
              onClick={() => setActiveTab('cae')}
              className={`w-full py-3 px-3 rounded-lg font-medium text-sm flex items-center justify-between transition ${activeTab === 'cae' ? 'bg-amber-500 text-white font-bold shadow-md' : 'bg-slate-800 text-white hover:bg-slate-700'}`}
            >
              <div className="flex items-center space-x-2">
                <HardHat className="w-5 h-5 text-amber-300" />
                <span>Módulo CAE</span>
              </div>
              <span className="text-[10px] bg-slate-900/40 px-1.5 py-0.5 rounded text-white font-mono">
                {currentUser.role === 'externo' ? 'UX' : currentUser.role === 'corporativo' ? 'UC' : 'SA'}
              </span>
            </button>
          </div>

          {currentUser.role === 'superadmin' && (
            <div className="pt-4 border-t border-slate-100 space-y-1">
              <span className="text-[10px] font-bold text-slate-400 uppercase tracking-wider px-2">Administración SA</span>
              
              <button 
                onClick={() => setActiveTab('excel')}
                className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2.5 transition ${activeTab === 'excel' ? 'bg-blue-50 text-blue-600 font-semibold' : 'text-slate-600 hover:bg-slate-50'}`}
              >
                <FileSpreadsheet className="w-4 h-4" />
                <span>Carga Masiva Excel</span>
              </button>

              <button 
                onClick={() => setActiveTab('users')}
                className={`w-full py-2.5 px-3 rounded-lg font-medium text-sm flex items-center space-x-2.5 transition ${activeTab === 'users' ? 'bg-blue-50 text-blue-600 font-semibold' : 'text-slate-600 hover:bg-slate-50'}`}
              >
                <Users className="w-4 h-4" />
                <span>Gestión Perfiles</span>
              </button>
            </div>
          )}
        </aside>

        {/* CONTENIDO PRINCIPAL */}
        <main className="flex-1">
          
          {/* TAB: CENTROS DE TRABAJO */}
          {activeTab === 'centres' && !currentSelectedCentre && (
            <div className="space-y-6">
              <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center gap-4">
                <div>
                  <h2 className="text-xl font-bold text-slate-800">Centros Asignados</h2>
                  <p className="text-sm text-slate-500">Selecciona un centro para ver sus documentos de Prevención.</p>
                </div>

                <div className="relative w-full sm:w-64">
                  <Search className="w-4 h-4 absolute left-3 top-3 text-slate-400" />
                  <input 
                    type="text"
                    placeholder="Buscar centro..."
                    className="w-full pl-9 pr-4 py-2 border border-slate-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                    value={searchTerm}
                    onChange={(e) => setSearchTerm(e.target.value)}
                  />
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                {accessibleCentres
                  .filter(c => c.name.toLowerCase().includes(searchTerm.toLowerCase()))
                  .map((centre) => {
                    const categoriesKeys = ['evaluacion_riesgos', 'informacion_riesgos', 'medidas_emergencia'];
                    const pendingCount = categoriesKeys.filter(k => {
                      const list = Array.isArray(centre.docs[k]) ? centre.docs[k] : [];
                      return list.length === 0;
                    }).length;

                    return (
                      <div 
                        key={centre.id}
                        onClick={() => setSelectedCentre(centre)}
                        className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm hover:shadow-md transition cursor-pointer flex flex-col justify-between"
                      >
                        <div className="space-y-3">
                          <div className="flex justify-between items-start">
                            <span className="text-xs font-semibold px-2.5 py-1 bg-slate-100 text-slate-600 rounded">
                              {centre.zone}
                            </span>
                            {pendingCount === 0 ? (
                              <span className="text-xs font-semibold px-2.5 py-1 bg-emerald-100 text-emerald-700 rounded-full flex items-center gap-1">
                                <CheckCircle2 className="w-3 h-3" /> Completo
                              </span>
                            ) : (
                              <span className="text-xs font-semibold px-2.5 py-1 bg-amber-100 text-amber-700 rounded-full flex items-center gap-1">
                                <AlertTriangle className="w-3 h-3" /> {pendingCount} Categ. sin docs
                              </span>
                            )}
                          </div>

                          <h3 className="text-lg font-bold text-slate-800">
                            {centre.name}
                          </h3>
                        </div>

                        <div className="mt-6 pt-4 border-t border-slate-100 flex justify-between items-center text-sm font-medium text-blue-600">
                          <span>Ver documentación</span>
                          <ExternalLink className="w-4 h-4" />
                        </div>
                      </div>
                    );
                  })}
              </div>
            </div>
          )}

          {/* DETALLE CENTRO SELECCIONADO */}
          {activeTab === 'centres' && currentSelectedCentre && (
            <div className="space-y-6">
              <button 
                onClick={() => setSelectedCentre(null)}
                className="text-sm font-medium text-slate-500 hover:text-slate-800 flex items-center gap-1"
              >
                ← Volver a centros
              </button>

              <div className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm space-y-2">
                <span className="text-xs font-semibold px-2.5 py-1 bg-slate-100 text-slate-600 rounded">
                  {currentSelectedCentre.zone}
                </span>
                <h2 className="text-2xl font-bold text-slate-800">{currentSelectedCentre.name}</h2>
                <p className="text-xs text-slate-400">
                  Usuarios con acceso: {currentSelectedCentre.users.join(', ')}
                </p>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                {[
                  { key: 'evaluacion_riesgos', label: 'Evaluación de Riesgos' },
                  { key: 'informacion_riesgos', label: 'Información de Riesgos' },
                  { key: 'medidas_emergencia', label: 'Medidas de Emergencia' }
                ].map(({ key, label }) => {
                  const docList = Array.isArray(currentSelectedCentre.docs[key]) ? currentSelectedCentre.docs[key] : [];
                  const isPresent = docList.length > 0;

                  return (
                    <div key={key} className="bg-white p-6 rounded-xl border border-slate-200 shadow-sm flex flex-col justify-between space-y-4">
                      <div>
                        <div className="flex justify-between items-center mb-3">
                          <span className="text-xs font-bold text-slate-400 uppercase tracking-wider">Categoría</span>
                          {isPresent ? (
                            <span className="text-xs font-bold px-2.5 py-1 bg-emerald-100 text-emerald-700 rounded-full flex items-center gap-1">
                              <CheckCircle2 className="w-3 h-3" /> {docList.length} Archivo(s)
                            </span>
                          ) : (
                            <span className="text-xs font-bold px-2.5 py-1 bg-rose-100 text-rose-700 rounded-full flex items-center gap-1">
                              <XCircle className="w-3 h-3" /> Pendiente
                            </span>
                          )}
                        </div>

                        <h3 className="text-base font-bold text-slate-800 mb-3">{label}</h3>

                        {isPresent ? (
                          <div className="space-y-2 max-h-48 overflow-y-auto pr-1">
                            {docList.map(doc => (
                              <a 
                                key={doc.id}
                                href={doc.link} 
                                target="_blank" 
                                rel="noreferrer"
                                className="p-2.5 bg-slate-50 hover:bg-blue-50 border border-slate-200 hover:border-blue-200 rounded-lg flex items-center justify-between text-xs transition group"
                              >
                                <div className="flex items-center space-x-2 truncate pr-2">
                                  <FileText className="w-4 h-4 text-blue-600 shrink-0" />
                                  <span className="font-medium text-slate-700 group-hover:text-blue-700 truncate">{doc.name}</span>
                                </div>
                                <ExternalLink className="w-3.5 h-3.5 text-slate-400 group-hover:text-blue-600 shrink-0" />
                              </a>
                            ))}
                          </div>
                        ) : (
                          <p className="text-xs text-rose-500 italic bg-rose-50 p-3 rounded border border-rose-100">
                            Sin documentos vinculados en esta categoría.
                          </p>
                        )}
                      </div>

                      <div className="space-y-2 pt-2 border-t border-slate-100">
                        {currentUser.role === 'superadmin' && (
                          <button 
                            onClick={() => {
                              setEditDocModal({
                                open: true,
                                centreId: currentSelectedCentre.id,
                                categoryKey: key,
                                categoryLabel: label
                              });
                              setNewLinkUrl('');
                              setNewLinkName('');
                            }}
                            className="w-full py-2 px-3 bg-slate-100 text-slate-700 font-medium text-xs rounded-lg hover:bg-slate-200 transition flex items-center justify-center gap-1.5"
                          >
                            <Plus className="w-3.5 h-3.5" />
                            <span>Gestionar / Añadir Enlaces</span>
                          </button>
                        )}
                      </div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* TAB: MÓDULO CAE (SEGÚN ROL) */}
          {activeTab === 'cae' && (
            <div className="space-y-6">
              <div className="bg-slate-800 text-white p-6 rounded-xl shadow-sm flex items-center justify-between">
                <div>
                  <div className="flex items-center space-x-2">
                    <HardHat className="w-6 h-6 text-amber-400" />
                    <h2 className="text-xl font-bold">Coordinación de Actividades Empresariales (CAE)</h2>
                  </div>
                  <p className="text-xs text-slate-300 mt-1">
                    {currentUser.role === 'superadmin' && 'Panel Global de Validación de Coordinación para Centros de Trabajo.'}
                    {currentUser.role === 'corporativo' && 'Consulta de empresas externas autorizadas para acceder a tus centros.'}
                    {currentUser.role === 'externo' && 'Recepción de documentación del centro y envío de coordinación de tu empresa.'}
                  </p>
                </div>

                <span className="hidden sm:block text-xs bg-slate-700 px-3 py-1.5 rounded-lg border border-slate-600 font-mono">
                  {currentUser.role === 'superadmin' && 'Modo: Validaciones (SA)'}
                  {currentUser.role === 'corporativo' && 'Modo: Consulta Centro (UC)'}
                  {currentUser.role === 'externo' && 'Modo: Empresa Externa (UX)'}
                </span>
              </div>

              {/* VISTA 1: USUARIO DE EMPRESA EXTERNA (UX) */}
              {currentUser.role === 'externo' && (
                <div className="space-y-6">
                  {accessibleCentres.map(centre => {
                    const record = caeRecords.find(r => r.centreId === centre.id && r.userEmail === currentUser.email);
                    const isRead = record ? record.docsRead : false;
                    const isSent = record ? record.docsSent : false;
                    const status = record ? record.status : 'pendiente_lectura';

                    return (
                      <div key={centre.id} className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-6">
                        <div className="flex flex-col sm:flex-row justify-between items-start sm:items-center border-b pb-4 gap-2">
                          <div>
                            <span className="text-xs font-semibold px-2 py-0.5 bg-slate-100 text-slate-600 rounded">
                              {centre.zone}
                            </span>
                            <h3 className="text-lg font-bold text-slate-800 mt-1">{centre.name}</h3>
                          </div>

                          <div>
                            {status === 'completado' && (
                              <span className="px-3 py-1 bg-emerald-100 text-emerald-800 text-xs font-bold rounded-full flex items-center gap-1.5">
                                <CheckCircle2 className="w-4 h-4 text-emerald-600" /> Coordinación Completada (Acceso Permitido)
                              </span>
                            )}
                            {status === 'documentos_enviados' && (
                              <span className="px-3 py-1 bg-amber-100 text-amber-800 text-xs font-bold rounded-full flex items-center gap-1.5">
                                <Clock className="w-4 h-4 text-amber-600" /> Pendiente de Validación por el SA
                              </span>
                            )}
                            {status === 'pendiente_lectura' && (
                              <span className="px-3 py-1 bg-slate-100 text-slate-700 text-xs font-bold rounded-full flex items-center gap-1.5">
                                <AlertTriangle className="w-4 h-4 text-slate-500" /> Pendiente de revisión
                              </span>
                            )}
                          </div>
                        </div>

                        {/* Paso 1: Consultar documentos del centro */}
                        <div className="space-y-3">
                          <h4 className="text-xs font-bold text-slate-500 uppercase tracking-wider">
                            1. Documentación e información del Centro para la Empresa Externa
                          </h4>

                          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
                            {[
                              { label: 'Evaluación de Riesgos', list: centre.docs.evaluacion_riesgos },
                              { label: 'Información de Riesgos', list: centre.docs.informacion_riesgos },
                              { label: 'Medidas de Emergencia', list: centre.docs.medidas_emergencia }
                            ].map((item, i) => (
                              <div key={i} className="p-3 bg-slate-50 rounded-lg border border-slate-200 text-xs">
                                <p className="font-bold text-slate-700 mb-1">{item.label}</p>
                                {item.list && item.list.length > 0 ? (
                                  item.list.map(d => (
                                    <a key={d.id} href={d.link} target="_blank" rel="noreferrer" className="text-blue-600 hover:underline flex items-center gap-1 truncate mt-0.5">
                                      <FileText className="w-3 h-3 shrink-0" /> {d.name}
                                    </a>
                                  ))
                                ) : (
                                  <span className="text-slate-400 italic">No disponible</span>
                                )}
                              </div>
                            ))}
                          </div>

                          <label className="flex items-center space-x-2 mt-3 p-3 bg-blue-50/50 rounded-lg border border-blue-100 cursor-pointer text-xs font-medium text-slate-700">
                            <input 
                              type="checkbox" 
                              checked={isRead}
                              disabled={status === 'completado'}
                              onChange={() => handleToggleReadDocs(centre.id)}
                              className="w-4 h-4 text-blue-600 rounded focus:ring-blue-500"
                            />
                            <span>He leído y recibido correctamente la información de PRL de este centro de trabajo.</span>
                          </label>
                        </div>

                        {/* Paso 2: Enviar Documentación y solicitar acceso */}
                        <div className="pt-4 border-t border-slate-100 flex flex-col sm:flex-row justify-between items-center gap-4">
                          <div className="text-xs text-slate-500">
                            <p className="font-semibold text-slate-700">2. Envío de documentación CAE de tu empresa</p>
                            <p>Envía la documentación requerida. El Superadministrador recibirá el aviso para liberar el acceso.</p>
                          </div>

                          <button 
                            disabled={!isRead || isSent}
                            onClick={() => handleSendCaeDocs(centre.id)}
                            className={`py-2.5 px-5 rounded-lg text-xs font-bold flex items-center gap-2 transition ${
                              isSent ? 'bg-emerald-600 text-white cursor-default' :
                              !isRead ? 'bg-slate-200 text-slate-400 cursor-not-allowed' :
                              'bg-blue-600 text-white hover:bg-blue-700 shadow-md'
                            }`}
                          >
                            <Send className="w-4 h-4" />
                            {isSent ? 'Documentación Enviada' : 'Enviar Documentación de Coordinación'}
                          </button>
                        </div>
                      </div>
                    );
                  })}
                </div>
              )}

              {/* VISTA 2: USUARIO CORPORATIVO (UC) */}
              {currentUser.role === 'corporativo' && (
                <div className="space-y-6">
                  <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <h3 className="text-base font-bold text-slate-800">
                      Empresas Externas con Acceso y Estado CAE en tus Centros
                    </h3>
                    <p className="text-xs text-slate-500">
                      Como usuario corporativo, consulta las contratas que han solicitado o completado la coordinación.
                    </p>

                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-xs text-slate-600">
                        <thead className="bg-slate-50 text-slate-500 font-semibold uppercase tracking-wider border-b">
                          <tr>
                            <th className="px-4 py-3">Centro</th>
                            <th className="px-4 py-3">Empresa Externa</th>
                            <th className="px-4 py-3">Correo Contacto</th>
                            <th className="px-4 py-3">Estado CAE</th>
                            <th className="px-4 py-3">Acceso al Centro</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {accessibleCentres.map(centre => {
                            const centreRecords = caeRecords.filter(r => r.centreId === centre.id);

                            if (centreRecords.length === 0) {
                              return (
                                <tr key={centre.id}>
                                  <td className="px-4 py-3 font-semibold text-slate-800">{centre.name}</td>
                                  <td colSpan="4" className="px-4 py-3 text-slate-400 italic">
                                    No hay empresas externas registradas para este centro.
                                  </td>
                                </tr>
                              );
                            }

                            return centreRecords.map(rec => (
                              <tr key={rec.id} className="hover:bg-slate-50">
                                <td className="px-4 py-3 font-semibold text-slate-800">{centre.name}</td>
                                <td className="px-4 py-3 font-medium text-slate-700">{rec.companyName}</td>
                                <td className="px-4 py-3">{rec.userEmail}</td>
                                <td className="px-4 py-3">
                                  {rec.status === 'completado' && (
                                    <span className="px-2.5 py-1 bg-emerald-100 text-emerald-700 font-bold rounded-full inline-flex items-center gap-1">
                                      <CheckCircle2 className="w-3 h-3" /> Validado por SA
                                    </span>
                                  )}
                                  {rec.status === 'documentos_enviados' && (
                                    <span className="px-2.5 py-1 bg-amber-100 text-amber-700 font-bold rounded-full inline-flex items-center gap-1">
                                      <Clock className="w-3 h-3" /> En revisión por SA
                                    </span>
                                  )}
                                  {rec.status === 'pendiente_lectura' && (
                                    <span className="px-2.5 py-1 bg-slate-100 text-slate-600 font-bold rounded-full inline-flex items-center gap-1">
                                      Pendiente
                                    </span>
                                  )}
                                </td>
                                <td className="px-4 py-3 font-bold">
                                  {rec.status === 'completado' ? (
                                    <span className="text-emerald-600 flex items-center gap-1">
                                      <Check className="w-4 h-4" /> PERMITIDO
                                    </span>
                                  ) : (
                                    <span className="text-rose-500 flex items-center gap-1">
                                      <XCircle className="w-4 h-4" /> DENEGADO / DENEGADO OK
                                    </span>
                                  )}
                                </td>
                              </tr>
                            ));
                          })}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}

              {/* VISTA 3: SUPERADMINISTRADOR (SA) */}
              {currentUser.role === 'superadmin' && (
                <div className="space-y-6">
                  <div className="bg-white rounded-xl border border-slate-200 p-6 shadow-sm space-y-4">
                    <div className="flex justify-between items-center border-b pb-3">
                      <div>
                        <h3 className="text-base font-bold text-slate-800">Validación de Solicitudes CAE y Liberación de Acceso</h3>
                        <p className="text-xs text-slate-500">
                          Revisa la documentación enviada por las empresas externas para pulsar OK y dejar libre su acceso.
                        </p>
                      </div>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-xs text-slate-600">
                        <thead className="bg-slate-50 text-slate-500 font-semibold uppercase tracking-wider border-b">
                          <tr>
                            <th className="px-4 py-3">Empresa Externa</th>
                            <th className="px-4 py-3">Centro Solicitado</th>
                            <th className="px-4 py-3">Lectura PRL</th>
                            <th className="px-4 py-3">Envío Doc.</th>
                            <th className="px-4 py-3">Estado</th>
                            <th className="px-4 py-3 text-right">Acción (Validación SA)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-100">
                          {caeRecords.length === 0 ? (
                            <tr>
                              <td colSpan="6" className="px-4 py-6 text-center text-slate-400 italic">
                                No hay registros de coordinaciones enviadas.
                              </td>
                            </tr>
                          ) : (
                            caeRecords.map(rec => {
                              const centre = centres.find(c => c.id === rec.centreId);

                              return (
                                <tr key={rec.id} className="hover:bg-slate-50">
                                  <td className="px-4 py-3">
                                    <p className="font-bold text-slate-800">{rec.companyName}</p>
                                    <p className="text-[10px] text-slate-400">{rec.userEmail}</p>
                                  </td>
                                  <td className="px-4 py-3 font-medium text-slate-700">{centre?.name}</td>
                                  <td className="px-4 py-3">
                                    {rec.docsRead ? (
                                      <span className="text-emerald-600 font-semibold">Sí (Confirmado)</span>
                                    ) : (
                                      <span className="text-slate-400">No</span>
                                    )}
                                  </td>
                                  <td className="px-4 py-3">
                                    {rec.docsSent ? (
                                      <span className="text-emerald-600 font-semibold">Sí (Recibido)</span>
                                    ) : (
                                      <span className="text-slate-400">No</span>
                                    )}
                                  </td>
                                  <td className="px-4 py-3">
                                    {rec.status === 'completado' && (
                                      <span className="px-2.5 py-1 bg-emerald-100 text-emerald-800 font-bold rounded-full">
                                        Acceso Libre
                                      </span>
                                    )}
                                    {rec.status === 'documentos_enviados' && (
                                      <span className="px-2.5 py-1 bg-amber-100 text-amber-800 font-bold rounded-full animate-pulse">
                                        Pendiente Validar
                                      </span>
                                    )}
                                    {rec.status === 'pendiente_lectura' && (
                                      <span className="px-2.5 py-1 bg-slate-100 text-slate-600 font-bold rounded-full">
                                        Pendiente Envío
                                      </span>
                                    )}
                                  </td>
                                  <td className="px-4 py-3 text-right">
                                    {rec.status === 'completado' ? (
                                      <span className="text-emerald-600 font-bold text-xs flex items-center justify-end gap-1">
                                        <Check className="w-4 h-4" /> Validado OK
                                      </span>
                                    ) : (
                                      <button 
                                        disabled={!rec.docsSent}
                                        onClick={() => handleApproveCae(rec.id)}
                                        className={`py-1.5 px-3 rounded-lg font-bold text-xs shadow transition flex items-center gap-1 ml-auto ${
                                          rec.docsSent 
                                            ? 'bg-emerald-600 text-white hover:bg-emerald-700' 
                                            : 'bg-slate-200 text-slate-400 cursor-not-allowed'
                                        }`}
                                      >
                                        <Check className="w-3.5 h-3.5" />
                                        Validar y Liberar Acceso (OK)
                                      </button>
                                    )}
                                  </td>
                                </tr>
                              );
                            })
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* TAB: DOCUMENTACIÓN GENERAL */}
          {activeTab === 'general' && (
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Documentación General de PRL</h2>
                <p className="text-sm text-slate-500">Procedimientos, normas, protocolos y plantillas comunes a todos los centros.</p>
              </div>

              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <table className="w-full text-left text-sm text-slate-600">
                  <thead className="bg-slate-50 text-xs font-semibold text-slate-500 uppercase tracking-wider border-b">
                    <tr>
                      <th className="px-6 py-3">Documento</th>
                      <th className="px-6 py-3">Categoría</th>
                      <th className="px-6 py-3">Fecha</th>
                      <th className="px-6 py-3 text-right">Acceso</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {generalDocs.map((doc) => (
                      <tr key={doc.id} className="hover:bg-slate-50">
                        <td className="px-6 py-4 font-medium text-slate-800">{doc.title}</td>
                        <td className="px-6 py-4">
                          <span className="px-2 py-1 bg-slate-100 text-slate-600 text-xs rounded">
                            {doc.category}
                          </span>
                        </td>
                        <td className="px-6 py-4 text-slate-400 text-xs">{doc.date}</td>
                        <td className="px-6 py-4 text-right">
                          <a 
                            href={doc.link} 
                            target="_blank" 
                            rel="noreferrer"
                            className="text-blue-600 hover:text-blue-800 text-xs font-semibold inline-flex items-center gap-1"
                          >
                            Ver archivo <ExternalLink className="w-3 h-3" />
                          </a>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {/* TAB: CARGA MASIVA (SÓLO SA) */}
          {activeTab === 'excel' && currentUser.role === 'superadmin' && (
            <div className="bg-white p-8 rounded-xl border border-slate-200 shadow-sm space-y-6">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Carga Masiva de Centros y Usuarios</h2>
                <p className="text-sm text-slate-500">Sube tu hoja de cálculo (.xlsx) para actualizar los centros y accesos.</p>
              </div>

              <div className="p-4 bg-slate-50 rounded-lg border border-slate-200 space-y-2 text-xs text-slate-600">
                <p className="font-bold text-slate-800">Estructura del Excel:</p>
                <ul className="list-disc pl-4 space-y-1">
                  <li><strong>Columna A:</strong> Nombre del Centro</li>
                  <li><strong>Columna B:</strong> Zona</li>
                  <li><strong>Columnas C, E, G, I, K:</strong> Usuario / Correo</li>
                  <li><strong>Columnas D, F, H, J, L:</strong> Contraseña del usuario</li>
                </ul>
              </div>

              <div className="border-2 border-dashed border-slate-300 rounded-xl p-8 text-center hover:border-blue-500 transition cursor-pointer bg-slate-50">
                <input 
                  type="file" 
                  accept=".xlsx, .xls"
                  onChange={handleFileUploadCentres}
                  className="hidden" 
                  id="excel-upload" 
                />
                <label htmlFor="excel-upload" className="cursor-pointer space-y-2 block">
                  <FileSpreadsheet className="w-10 h-10 text-emerald-600 mx-auto" />
                  <span className="block text-sm font-medium text-slate-700">Seleccionar archivo Excel</span>
                </label>
              </div>
            </div>
          )}

          {/* TAB: GESTIÓN DE PERFILES (SÓLO SA) */}
          {activeTab === 'users' && currentUser.role === 'superadmin' && (
            <div className="space-y-6">
              <div>
                <h2 className="text-xl font-bold text-slate-800">Usuarios y Roles Registrados</h2>
                <p className="text-sm text-slate-500">Listado de credenciales, roles e identificación de usuarios.</p>
              </div>

              <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
                <table className="w-full text-left text-sm text-slate-600">
                  <thead className="bg-slate-50 text-xs font-semibold text-slate-500 uppercase tracking-wider border-b">
                    <tr>
                      <th className="px-6 py-3">Nombre / Empresa</th>
                      <th className="px-6 py-3">Usuario / Correo</th>
                      <th className="px-6 py-3">Contraseña</th>
                      <th className="px-6 py-3">Rol</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {users.map((u) => (
                      <tr key={u.id} className="hover:bg-slate-50">
                        <td className="px-6 py-4 font-medium text-slate-800">{u.name}</td>
                        <td className="px-6 py-4">{u.email}</td>
                        <td className="px-6 py-4 font-mono text-xs text-slate-600">{u.code}</td>
                        <td className="px-6 py-4">
                          <span className={`px-2 py-1 text-xs font-semibold rounded ${
                            u.role === 'superadmin' ? 'bg-purple-100 text-purple-700' :
                            u.role === 'corporativo' ? 'bg-blue-100 text-blue-700' : 'bg-emerald-100 text-emerald-700'
                          }`}>
                            {u.role === 'superadmin' ? 'Superadmin (SA)' :
                             u.role === 'corporativo' ? 'Usuario Corporativo (UC)' : 'Empresa Externa (UX)'}
                          </span>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

        </main>
      </div>

      {/* Modal para Enlaces de SharePoint por Categoría */}
      {editDocModal.open && (
        <div className="fixed inset-0 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4 z-50">
          <div className="bg-white rounded-xl shadow-xl max-w-lg w-full p-6 space-y-6 max-h-[90vh] flex flex-col">
            <div className="flex justify-between items-center border-b pb-3">
              <div>
                <h3 className="text-lg font-bold text-slate-800">Gestionar Enlaces SharePoint</h3>
                <p className="text-xs text-slate-500">{editDocModal.categoryLabel}</p>
              </div>
              <button 
                onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, categoryLabel: '' })}
                className="text-slate-400 hover:text-slate-600 text-xl font-bold"
              >
                &times;
              </button>
            </div>

            <form onSubmit={handleAddLink} className="space-y-3 bg-slate-50 p-4 rounded-lg border border-slate-200">
              <span className="text-xs font-bold text-slate-700 uppercase block">Añadir nuevo enlace</span>
              
              <div>
                <label className="block text-xs text-slate-500 mb-1">URL / Enlace de SharePoint</label>
                <div className="relative">
                  <LinkIcon className="w-4 h-4 absolute left-3 top-2.5 text-slate-400" />
                  <input 
                    type="url" 
                    required
                    className="w-full pl-9 pr-3 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                    placeholder="https://empresa.sharepoint.com/..."
                    value={newLinkUrl}
                    onChange={(e) => handleUrlChange(e.target.value)}
                  />
                </div>
              </div>

              <div>
                <label className="block text-xs text-slate-500 mb-1">Nombre del Archivo (Autodetectado o personalizado)</label>
                <input 
                  type="text" 
                  className="w-full px-3 py-2 border border-slate-300 rounded-lg text-xs focus:outline-none focus:ring-2 focus:ring-blue-500"
                  placeholder="Ej. Evaluacion_Mestalla_2026.pdf"
                  value={newLinkName}
                  onChange={(e) => setNewLinkName(e.target.value)}
                />
              </div>

              <button 
                type="submit"
                className="w-full py-2 bg-blue-600 text-white font-medium text-xs rounded-lg hover:bg-blue-700 transition shadow flex items-center justify-center gap-1"
              >
                <Plus className="w-4 h-4" />
                Guardar Enlace
              </button>
            </form>

            <div className="flex-1 overflow-y-auto space-y-2">
              <span className="text-xs font-bold text-slate-700 uppercase block">Enlaces Guardados</span>
              
              {(() => {
                const centre = centres.find(c => c.id === editDocModal.centreId);
                const docList = centre && Array.isArray(centre.docs[editDocModal.categoryKey]) 
                  ? centre.docs[editDocModal.categoryKey] 
                  : [];

                if (docList.length === 0) {
                  return (
                    <p className="text-xs text-slate-400 text-center py-4 italic border border-dashed rounded-lg">
                      No hay enlaces guardados en esta categoría.
                    </p>
                  );
                }

                return docList.map(doc => (
                  <div key={doc.id} className="flex items-center justify-between p-3 bg-white border border-slate-200 rounded-lg text-xs shadow-sm">
                    <div className="flex items-center space-x-2 truncate pr-2">
                      <FileText className="w-4 h-4 text-blue-600 shrink-0" />
                      <div className="truncate">
                        <p className="font-semibold text-slate-800 truncate">{doc.name}</p>
                        <a href={doc.link} target="_blank" rel="noreferrer" className="text-slate-400 hover:text-blue-600 truncate block text-[10px]">
                          {doc.link}
                        </a>
                      </div>
                    </div>

                    <button 
                      onClick={() => handleDeleteLink(doc.id)}
                      className="p-1.5 text-slate-400 hover:text-rose-600 hover:bg-rose-50 rounded transition shrink-0"
                      title="Eliminar enlace"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  </div>
                ));
              })()}
            </div>

            <div className="flex justify-end pt-3 border-t">
              <button 
                onClick={() => setEditDocModal({ open: false, centreId: null, categoryKey: null, categoryLabel: '' })}
                className="px-4 py-2 text-xs font-medium bg-slate-800 text-white rounded-lg hover:bg-slate-700 transition"
              >
                Cerrar
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}
