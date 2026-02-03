// LIEReport.tsx
import React, {
    useEffect,
    useMemo,
    useRef,
    useState,
    useCallback,
    memo,
    useTransition,
} from 'react';
import { useParams } from 'react-router-dom';
import {
    DndContext,
    DragEndEvent,
    DragOverEvent,
    DragStartEvent,
    PointerSensor,
    UniqueIdentifier,
    closestCenter,
    useSensor,
    useSensors,
} from '@dnd-kit/core';
import {
    SortableContext,
    useSortable,
    arrayMove,
    verticalListSortingStrategy,
} from '@dnd-kit/sortable';
import { CSS } from '@dnd-kit/utilities';
import {
    Box,
    Paper,
    Typography,
    Chip,
    IconButton,
    Collapse,
    Divider,
    Stack,
    List,
    ListItem,
    ListItemIcon,
    Checkbox,
    ListItemText,
    Button,
    TextField,
    Tooltip,
    Dialog,
    DialogTitle,
    DialogContent,
    DialogActions,
    Menu,
    MenuItem,
    Select,
    FormControl,
    InputLabel,
    Alert,
    CircularProgress,
    Drawer,
    Snackbar,
} from '@mui/material';
import {
    DragIndicatorRounded,
    ExpandMore,
    EditOutlined,
    SaveOutlined,
    CloseOutlined,
    AddOutlined,
    DeleteOutline,
    Menu as MenuIcon,
    Download,
    PictureAsPdf,
    Description,
    ArrowDropDown,
    Edit,
    Delete,
    CommentOutlined,
    AutoAwesome,
    Rotate90DegreesCcw,
    Print,
    Pageview,
    CheckCircle,
    Cancel,
    ThumbUp,
    ThumbDown
} from '@mui/icons-material';

import { usePrint } from '../hooks/usePrint';
import config from '../config';
import { useAuth } from '../context/AuthContext';
import { permissionService } from '../Roles/permissionService';
import logo from '../logo.png';
import ReportVersionProgress from "../pages/ReportVersionProgress";
import type { ReportStatusLog } from "../pages/ReportVersionProgress";

// Dynamically import docx and file-saver to avoid bundle issues
import { Document, Packer, Paragraph, Table, TableCell, TableRow, HeadingLevel, AlignmentType, TextRun, WidthType, PageBreak, PageOrientation, SectionType } from 'docx';
import { saveAs } from 'file-saver';

/* ------------------------------------------------- Constants ------------------------------------------------- */
const DEFAULT_IMAGES = [
    "https://placehold.co/400x400",
    "https://placehold.co/400x400",
    "https://placehold.co/400x400",
    "https://placehold.co/400x400",
    "https://placehold.co/400x400",
    "https://placehold.co/400x400",
    "https://placehold.co/400x400",
    "https://placehold.co/400x400",
];

const STAGES = [
    { stage: "Draft", label: "Draft Report", key: "draft_report" },
    { stage: "Stable", label: "Stable Report", key: "stable_report" },
    { stage: "Approval", label: "Report Approval", key: "report_approval" },
    { stage: "Final", label: "Report Final Review", key: "report_final_review" }
];

/* ------------------------------------------------- Types ------------------------------------------------- */
interface UserDataInfo {
    id: string;
    email: string;
    name: string;
    phone: string;
    role_id: string;
}

interface RoleDataInfo {
    id: string;
    name: string;
    permissions: string;
    view_all: string;
    manage_roles: string;
    manage_users: string;
    created_at: string;
    updated_at: string;
    deleted_at: string;
}

interface Project {
    id?: string;
    name?: string;
    client_id?: string;
    bank_id?: string;
    industry_id?: string;
    scope?: string;
    created_by?: string;
    assigned_analyst_id?: string;
    status?: string;
    created_at?: string;
    updated_at?: string;
    deleted_at?: string;
    project_type?: string;
    cover_img?: string;
    version_id?: string;
}

interface Banks {
    id?: string;
    name?: string;
    logo_url?: string;
}

interface Clients {
    id?: string;
    email?: string;
    name?: string;
    phone?: string;
    company_name?: string;
    created_at?: string;
    updated_at?: string;
    deleted_at?: string;
}

interface Industry {
    id: string;
    name: string;
    checklist_count?: number;
}

interface SectionTemplates {
    id?: string;
    industry_id?: string;
    name?: string;
    default_content?: any;
    is_mandatory?: boolean;
    order_suggestion?: number;
    section_type?: string;
    created_by?: string;
    created_at?: string;
    updated_at?: string;
    deleted_at?: string;
    project_type?: string;
    checked: boolean;
}

interface SubSectionTemplates {
    id?: string;
    section_template_id?: string;
    name?: string;
    default_content?: any;
    is_mandatory?: boolean;
    order_suggestion?: number;
    created_by?: string;
    created_at?: string;
    updated_at?: string;
    deleted_at?: string;
    checked: boolean;
    project_type?: string;
}

interface ReportVersions {
    id?: string;
    project_id?: string;
    version_number?: string;
    stage?: string;
    status?: string;
    is_draft?: string;
    pdf_url?: string;
    approved_by?: string;
    approved_at?: string;
    created_by?: string;
    created_at?: string;
    updated_at?: string;
    deleted_at?: string;
}

interface ProjectReportSectionsP {
    id: string;
    report_version_id: string;
    section_template_id: string;
    subsection_template_id: string;
    name: string;
    content: any;
    order_index: string;
    order_sub_index: string;
    is_deleted: string;
    edited_by: string;
    created_at: string;
    updated_at: string;
    deleted_at: string;
    project_id: string;
    checked: boolean;
}

interface Task {
    id: string;
    main_id: string | null;
    label: string;
    checked?: boolean;
    note?: string;
    olgpt?: string;
    order_sub_index?: number;
    content?: any;
}

interface Group {
    id: string;
    main_id?: string | null;
    title: string;
    items: Task[];
    checked?: boolean;
    expanded?: boolean;
    order_index?: string;
    content?: any;
}

interface Comment {
    id: string;
    comment: string;
    grammar_checked: boolean;
    user_id: string;
    entity_id: string;
    entity_type: string;
}

interface ExtractedData {
    [key: string]: any;
}

interface ChecklistItem {
    item: string;
    status: 'Uploaded' | 'Pending' | string;
    url?: string;
    document_id?: string | null;
    extracted_data?: ExtractedData | null;
}

interface ChecklistEntry {
    id: string;
    project_id: string;
    template_id: string;
    items: ChecklistItem[];
    sent_at?: string;
    completed_at?: string | null;
    created_by?: string;
    created_at?: string;
    updated_at?: string | null;
}

interface OlgptParagraph {
    content_type: 'paragraph';
    content: string;
}

interface OlgptTableHorizontal {
    content_type: 'table_horizontal';
    headers: string[];
    rows: string[][];
}

interface OgField {
    field_name: string;
    value: string;
}

interface OlgptTableVertical {
    content_type: 'table_vertical';
    content: OgField[];
}

// Add new interfaces for processing jobs
interface TocItem {
    query: string;
    status: 'pending' | 'approved' | 'rejected' | string;
    section_id: string;
    section_name: string;
    subsection_id: string;
    content_blocks: any[];
    section_number: string;
    subsection_name: string;
    subsection_number: string;
    project_report_sections_id: string;
}

interface ReportJson {
    toc: TocItem[];
}

interface ProcessingJobLog {
    id: string;
    project_id: string;
    olgpt_job_id: string;
    status: 'processing' | 'completed' | string;
    report_version_id: string;
    generation_type: string;
    created_at: string;
    report_json?: ReportJson;
    tocItems?: TocItem[];
}

type OlgptAny = OlgptParagraph | OlgptTableHorizontal | OlgptTableVertical;

/* ------------------------------------------------- Generation Query Dialog Component ------------------------------------------------- */
interface GenerationQueryDialogProps {
    open: boolean;
    onClose: () => void;
    onSubmit: (query: string) => void;
    title?: string;
    loading?: boolean;
}

const GenerationQueryDialog: React.FC<GenerationQueryDialogProps> = memo(({
    open,
    onClose,
    onSubmit,
    title = "Add Custom Instructions",
    loading = false
}) => {
    const [query, setQuery] = useState("");

    const handleSubmit = () => {
        if (query.trim()) {
            onSubmit(query.trim());
            setQuery("");
        }
    };

    const handleKeyPress = (e: React.KeyboardEvent) => {
        if (e.key === 'Enter' && !e.shiftKey) {
            e.preventDefault();
            handleSubmit();
        }
    };

    return (
        <Dialog open={open} onClose={onClose} maxWidth="sm" fullWidth>
            <DialogTitle>{title}</DialogTitle>
            <DialogContent>
                <Typography variant="body2" color="text.secondary" gutterBottom>
                    Add specific instructions or requirements for content generation:
                </Typography>
                <TextField
                    autoFocus
                    margin="dense"
                    label="Your instructions"
                    placeholder="e.g., Make the content more technical, focus on financial metrics, include recent market trends..."
                    multiline
                    rows={4}
                    fullWidth
                    value={query}
                    onChange={(e) => setQuery(e.target.value)}
                    onKeyPress={handleKeyPress}
                    disabled={loading}
                />
                <Typography variant="caption" color="text.secondary" sx={{ display: 'block', mt: 1 }}>
                    Press Enter to submit, Shift+Enter for new line
                </Typography>
            </DialogContent>
            <DialogActions>
                <Button onClick={onClose} disabled={loading}>Cancel</Button>
                <Button
                    onClick={handleSubmit}
                    variant="contained"
                    disabled={!query.trim() || loading}
                    sx={{ backgroundColor: '#032F5D' }}
                >
                    {loading ? 'Generating...' : 'Generate with Instructions'}
                </Button>
            </DialogActions>
        </Dialog>
    );
});

/* ------------------------------------------------- Custom Hooks ------------------------------------------------- */
const useSafeFetch = () => {
    const safeFetch = useCallback(async (url: string): Promise<any> => {
        const res = await fetch(`${config.backendURL}${url}`, {
            credentials: "include",
        });

        if (!res.ok) {
            const txt = await res.text();
            throw new Error(JSON.stringify({
                status: res.status,
                message: txt || res.statusText,
                url: `${config.backendURL}${url}`
            }));
        }

        const json = await res.json();
        return json.data ?? json;
    }, []);

    return safeFetch;
};

const useProjectData = (projectId: string) => {
    const [project, setProject] = useState<Project | null>(null);
    const [bank, setBank] = useState<Banks | null>(null);
    const [industry, setIndustry] = useState<Industry | null>(null);
    const [client, setClient] = useState<Clients[]>([]);
    const [loading, setLoading] = useState(true);
    const [error, setError] = useState("");
    const safeFetch = useSafeFetch();

    const load = useCallback(async () => {
        if (!projectId) {
            setError("No project id in the URL");
            return;
        }

        setLoading(true);
        setError("");

        try {
            const proj = await safeFetch(`/api/projects/${projectId}`);
            setProject(proj);

            const [clientP, industryP, bankP] = await Promise.allSettled([
                safeFetch(`/api/clients/${proj.client_id}`),
                safeFetch(`/api/industries/${proj.industry_id}`),
                safeFetch(`/api/banks/${proj.bank_id}`),
            ]);

            if (clientP.status === "fulfilled") setClient(clientP.value);
            if (industryP.status === "fulfilled") setIndustry(industryP.value);
            if (bankP.status === "fulfilled") setBank(bankP.value);
        } catch (e: any) {
            console.error(e);
            setError("Failed to load core project data. Please check the network connection and project API.");
        } finally {
            setLoading(false);
        }
    }, [projectId, safeFetch]);

    useEffect(() => {
        load();
    }, [load]);

    return {
        project,
        bank,
        industry,
        client,
        loading,
        error,
        reload: load,
    };
};

const useReportData = (project: Project | null) => {
    const [allSection, setAllSection] = useState<SectionTemplates[]>([]);
    const [allSubSection, setAllSubSection] = useState<SubSectionTemplates[]>([]);
    const [reportV, setReportV] = useState<ReportVersions | null>(null);
    const [projectRS, setProjectRS] = useState<ProjectReportSectionsP[]>([]);
    const [checklists, setChecklists] = useState<ChecklistEntry[]>([]);
    const safeFetch = useSafeFetch();

    const load = useCallback(async () => {
        if (!project) return;

        try {
            const [
                checklistsP,
                sectionTemplatesP,
                subsectionTemplatesP,
                reportVersionsP,
                projectReportSectionsP
            ] = await Promise.allSettled([
                safeFetch(`/api/checklists/`),
                safeFetch(`/api/section-templates`),
                safeFetch(`/api/subsection-templates`),
                safeFetch(`/api/report-versions/${project.version_id}`),
                safeFetch(`/api/project-report-sections`),
            ]);

            if (checklistsP.status === "fulfilled") setChecklists(checklistsP.value);
            if (sectionTemplatesP.status === "fulfilled") setAllSection(sectionTemplatesP.value);
            if (subsectionTemplatesP.status === "fulfilled") setAllSubSection(subsectionTemplatesP.value);
            if (reportVersionsP.status === "fulfilled") setReportV(reportVersionsP.value);
            if (projectReportSectionsP.status === "fulfilled") setProjectRS(projectReportSectionsP.value);
        } catch (e) {
            console.error("Failed to load report data:", e);
        }
    }, [project, safeFetch]);

    useEffect(() => {
        load();
    }, [load]);

    return {
        allSection,
        allSubSection,
        reportV,
        projectRS,
        checklists,
        reload: load,
        setAllSection,
        setAllSubSection,
        setReportV,
        setProjectRS,
    };
};

/* ------------------------------------------------- Comment Modal Component ------------------------------------------------- */
interface CommentModalProps {
    open: boolean;
    onClose: () => void;
    entityType: string;
    entityId: string;
    safeFetch: (url: string) => Promise<any>;
    selectedEntity: { type: string; id: string };
    grammarCheck?: (text: string) => Promise<string>;
    loading?: boolean;
    setLoading?: (loading: boolean) => void;
}

const CommentModal = memo(({
    open,
    onClose,
    entityType,
    entityId,
    safeFetch,
    selectedEntity,
    grammarCheck,
    loading,
    setLoading
}: CommentModalProps) => {
    const { userId } = useAuth();
    const [editCommentId, setEditCommentId] = useState<string | null>(null);
    const [userDetail, setUserDetail] = useState<UserDataInfo[]>([]);
    const [comments, setComments] = useState<Comment[]>([]);
    const [newComment, setNewComment] = useState("");

    useEffect(() => {
        if (open) {
            fetchComments();
            fetchUserInfo();
        }
    }, [open]);

    const fetchComments = async () => {
        const data = await safeFetch(`/api/comments`);
        setComments(data);
    };

    const fetchUserInfo = async () => {
        const data = await safeFetch(`/api/users`);
        setUserDetail(Array.isArray(data) ? data : []);
    };

    const updateComment = async () => {
        if (newComment.trim() === "" || !editCommentId) return;

        await fetch(`${config.backendURL}/api/comments/${editCommentId}`, {
            method: "PUT",
            credentials: "include",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                entity_type: entityType,
                entity_id: entityId,
                comment: newComment,
                grammar_checked: false
            })
        });

        setNewComment("");
        setEditCommentId(null);
        fetchComments();
    };

    const createComment = async () => {
        if (newComment.trim() === "") return;

        await fetch(`${config.backendURL}/api/comments`, {
            method: "POST",
            credentials: "include",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                entity_type: entityType,
                entity_id: entityId,
                comment: newComment,
                grammar_checked: false
            })
        });

        setNewComment("");
        fetchComments();
    };

    const handleGrammarCheck = async (commentId: string, commentText: string) => {
        if (!grammarCheck || !setLoading) return;

        setLoading(true);
        try {
            const grammarResult = await grammarCheck(commentText);
            setEditCommentId(commentId);
            setNewComment(grammarResult);
        } finally {
            setLoading(false);
        }
    };

    const handleDeleteComment = async (commentId: string) => {
        try {
            await fetch(`${config.backendURL}/api/comments/${commentId}`, {
                method: 'DELETE',
                credentials: 'include'
            });
            fetchComments();
        } catch (e) {
            console.error(e);
        }
    };

    if (!open) return null;

    const filteredComments = comments.filter(
        c => c.entity_id === selectedEntity.id && c.entity_type === selectedEntity.type
    );

    return (
        <div className="fixed inset-0 bg-black/40 flex items-center justify-center z-[9999]">
            <div className="bg-white rounded-lg p-4 w-3/5 shadow-xl">
                <h2 className="text-lg font-bold mb-3">Comments</h2>

                <div className="max-h-60 overflow-auto border p-2 rounded">
                    {filteredComments.map(c => {
                        const user = userDetail.find(u => u.id === c.user_id);
                        const initials = user?.name
                            ? user.name.split(' ').map(p => p[0]).slice(0, 2).join('').toUpperCase()
                            : '?';

                        return (
                            <div key={c.id} className="py-2">
                                <Paper variant="outlined" className="p-3">
                                    <Stack direction="row" spacing={2} alignItems="flex-start">
                                        <Box
                                            sx={{
                                                width: 40,
                                                height: 40,
                                                borderRadius: '50%',
                                                bgcolor: 'primary.main',
                                                color: 'common.white',
                                                display: 'flex',
                                                alignItems: 'center',
                                                justifyContent: 'center',
                                                fontWeight: 700,
                                                fontSize: 14,
                                                flexShrink: 0,
                                            }}
                                            aria-hidden
                                        >
                                            {initials}
                                        </Box>

                                        <Box sx={{ flex: 1 }}>
                                            <Stack direction="row" alignItems="center" justifyContent="space-between">
                                                <Typography variant="subtitle2" sx={{ fontWeight: 700 }}>
                                                    {user?.name || "Loading..."}
                                                </Typography>

                                                <Stack direction="row" spacing={1} alignItems="center">
                                                    {c.grammar_checked ? (
                                                        <Chip size="small" label="Grammar checked" color="success" />
                                                    ) : (
                                                        <Chip size="small" label="Pending grammar" variant="outlined" />
                                                    )}
                                                    {c.user_id === String(userId) && (
                                                        <>
                                                            <IconButton
                                                                size="small"
                                                                onClick={() => {
                                                                    setEditCommentId(c.id);
                                                                    setNewComment(c.comment);
                                                                }}
                                                                aria-label="edit comment"
                                                            >
                                                                <Edit fontSize="small" />
                                                            </IconButton>

                                                            <IconButton
                                                                size="small"
                                                                onClick={() => handleDeleteComment(c.id)}
                                                                aria-label="delete comment"
                                                            >
                                                                <Delete fontSize="small" />
                                                            </IconButton>
                                                        </>
                                                    )}
                                                </Stack>
                                            </Stack>

                                            <Divider sx={{ my: 1 }} />

                                            <Typography variant="body2" sx={{ whiteSpace: 'pre-wrap', color: 'text.primary' }}>
                                                {c.comment}
                                            </Typography>

                                            {c.user_id === String(userId) && grammarCheck && setLoading && (
                                                <Stack direction="row" spacing={1} sx={{ mt: 1 }} className="print:hidden">
                                                    <Button
                                                        size="small"
                                                        onClick={() => handleGrammarCheck(c.id, c.comment)}
                                                        disabled={loading}
                                                    >
                                                        Request Grammar Check
                                                    </Button>
                                                </Stack>
                                            )}
                                        </Box>
                                    </Stack>
                                </Paper>
                            </div>
                        );
                    })}

                    {filteredComments.length === 0 && (
                        <div className="text-gray-400 text-sm text-center py-4">
                            No comments yet.
                        </div>
                    )}
                </div>

                <TextField
                    className="mt-3"
                    placeholder="Write a comment..."
                    value={newComment}
                    onChange={(e) => setNewComment(e.target.value)}
                    variant="outlined"
                    multiline
                    rows={4}
                    fullWidth
                />

                <div className="flex justify-end gap-2 mt-3">
                    <Button variant="outlined" onClick={() => {
                        setEditCommentId(null);
                        setNewComment("");
                        onClose();
                    }}>
                        Close
                    </Button>

                    <Button
                        variant="contained"
                        color="primary"
                        onClick={editCommentId ? updateComment : createComment}
                        disabled={!newComment.trim()}
                    >
                        {editCommentId ? "Update" : "Add"}
                    </Button>
                </div>
            </div>
        </div>
    );
});


/* ------------------------------------------------- Section Comments Component ------------------------------------------------- */
const SectionComments = memo(({
    sectionId,
    safeFetch,
}: {
    sectionId: string;
    safeFetch: (url: string) => Promise<any>;
}) => {
    const [comments, setComments] = useState<Comment[]>([]);
    const [users, setUsers] = useState<UserDataInfo[]>([]);
    const [commentsReloadTick, setCommentsReloadTick] = useState(0);

    useEffect(() => {
        const load = async () => {
            const [commentsData, usersData] = await Promise.all([
                safeFetch("/api/comments"),
                safeFetch("/api/users"),
            ]);

            setComments(
                commentsData.filter(
                    (c: Comment) =>
                        c.entity_type === "section" &&
                        c.entity_id === sectionId
                )
            );

            setCommentsReloadTick(t => t + 1);
            setUsers(Array.isArray(usersData) ? usersData : []);
        };

        load();
    }, [sectionId, safeFetch, commentsReloadTick]);

    return (
        <Box sx={{ width: "100%", flex: 1, overflow: 'auto' }}>
            {comments.length === 0 ? (
                <Box sx={{
                    textAlign: 'center',
                    py: 4,
                    color: 'text.secondary',
                    fontStyle: 'italic'
                }}>
                    No comments yet for this section.
                </Box>
            ) : (
                <Stack spacing={2}>
                    {comments.map(c => {
                        const user = users.find(u => u.id === c.user_id);

                        return (
                            <Box
                                key={c.id}
                                sx={{
                                    display: "flex",
                                    gap: 1.5,
                                    p: 2,
                                    backgroundColor: "#f8fafc",
                                    border: "1px solid #e2e8f0",
                                    borderRadius: "8px",
                                    boxShadow: '0 1px 3px rgba(0,0,0,0.05)'
                                }}
                            >
                                <Box
                                    sx={{
                                        width: "4px",
                                        borderRadius: "2px",
                                        backgroundColor: "#1e40af",
                                        flexShrink: 0,
                                    }}
                                />

                                <Box sx={{ flex: 1 }}>
                                    <Typography
                                        sx={{
                                            fontSize: "0.9rem",
                                            lineHeight: 1.7,
                                            color: "#1e293b",
                                            whiteSpace: "pre-wrap",
                                        }}
                                    >
                                        {c.comment}
                                    </Typography>

                                    <Box sx={{
                                        display: 'flex',
                                        justifyContent: 'space-between',
                                        alignItems: 'center',
                                        mt: 1.5,
                                        pt: 1,
                                        borderTop: '1px solid #e2e8f0'
                                    }}>
                                        <Typography
                                            sx={{
                                                fontSize: "0.75rem",
                                                color: "#64748b",
                                                fontWeight: 500
                                            }}
                                        >
                                            {user?.name || "Unknown User"}
                                        </Typography>

                                        {c.grammar_checked && (
                                            <Chip
                                                size="small"
                                                label="Grammar Checked"
                                                color="success"
                                                sx={{ height: 20, fontSize: '0.65rem' }}
                                            />
                                        )}
                                    </Box>
                                </Box>
                            </Box>
                        );
                    })}
                </Stack>
            )}
        </Box>
    );
});

/* ------------------------------------------------- Subsection Comments Component ------------------------------------------------- */
const SubsectionComments = memo(({
    subsectionId,
    safeFetch,
}: {
    subsectionId: string;
    safeFetch: (url: string) => Promise<any>;
}) => {
    const [comments, setComments] = useState<Comment[]>([]);
    const [users, setUsers] = useState<UserDataInfo[]>([]);
    const [commentsReloadTick, setCommentsReloadTick] = useState(0);

    useEffect(() => {
        const load = async () => {
            const [commentsData, usersData] = await Promise.all([
                safeFetch("/api/comments"),
                safeFetch("/api/users"),
            ]);

            setComments(
                commentsData.filter(
                    (c: Comment) =>
                        c.entity_type === "subsection" &&
                        c.entity_id === subsectionId
                )
            );
            setCommentsReloadTick(t => t + 1);
            setUsers(Array.isArray(usersData) ? usersData : []);
        };

        load();
    }, [subsectionId, safeFetch, commentsReloadTick]);

    if (!comments.length) return null;

    return (
        <Box sx={{ width: "100%", mt: 1, border: "1px solid #e5e7eb", borderRadius: "6px", p: 1.5 }}>
            <Typography variant="subtitle2" sx={{ fontWeight: 'bold', mb: 1, color: '#374151' }}>
                Comments
            </Typography>
            <Stack spacing={1.5}>
                {comments.map(c => {
                    const user = users.find(u => u.id === c.user_id);

                    return (
                        <Box
                            key={c.id}
                            sx={{
                                display: "flex",
                                gap: 1,
                                p: 1.5,
                                backgroundColor: "#f9fafb",
                                border: "1px solid #e5e7eb",
                                borderRadius: "6px",
                            }}
                        >
                            <Box
                                sx={{
                                    width: "4px",
                                    borderRadius: "2px",
                                    backgroundColor: "#2563eb",
                                    flexShrink: 0,
                                }}
                            />

                            <Box sx={{ flex: 1 }}>
                                <Typography
                                    sx={{
                                        fontSize: "0.8rem",
                                        lineHeight: 1.6,
                                        color: "#111827",
                                        whiteSpace: "pre-wrap",
                                    }}
                                >
                                    {c.comment}
                                </Typography>

                                <Typography
                                    sx={{
                                        fontSize: "0.65rem",
                                        color: "#6b7280",
                                        mt: 0.75,
                                    }}
                                >
                                    {user?.name || "Unknown"}
                                </Typography>
                            </Box>
                        </Box>
                    );
                })}
            </Stack>
        </Box>
    );
});

/* ------------------------------------------------- Utility Functions ------------------------------------------------- */
const formatMonthYear = (dateInput?: string | Date): string => {
    if (!dateInput) return "";
    const date = typeof dateInput === "string" ? new Date(dateInput) : dateInput;
    if (isNaN(date.getTime())) return "";
    return date.toLocaleString("en-US", { month: "long", year: "numeric" });
};

const normalizeContent = (content: any): any => {
    if (content === undefined || content === null) {
        return null;
    }

    if (typeof content === 'string') {
        try {
            const parsed = JSON.parse(content);
            return normalizeContent(parsed);
        } catch (e) {
            return [];
        }
    }

    if (Array.isArray(content)) {
        return content.map((item: any) => {
            if (item && typeof item === 'object') {
                const cleaned = { ...item };
                if (cleaned.content && typeof cleaned.content === 'string') {
                    const trimmed = cleaned.content.trim();
                    if ((trimmed.startsWith('"') && trimmed.endsWith('"') && trimmed.length > 2) ||
                        (trimmed.startsWith("'") && trimmed.endsWith("'") && trimmed.length > 2)) {
                        try {
                            const parsed = JSON.parse(cleaned.content);
                            cleaned.content = typeof parsed === 'string' ? parsed : cleaned.content;
                        } catch (e) {
                            // Keep as is
                        }
                    }
                }
                return cleaned;
            }
            return item;
        }).filter((item: any) => item !== null && item !== undefined);
    }

    if (content && typeof content === 'object') {
        return content;
    }

    return [];
};

const simplifyReportData = (data: any): [any[], any[]] => {
    if (!Array.isArray(data)) return [[], []];

    const section_temps: any[] = [];
    const sub_section_temps: any[] = [];

    data.forEach((group: any) => {
        section_temps.push({
            section_template_id: group.id,
            order_index: group.order_index,
        });

        if (Array.isArray(group.items)) {
            group.items.forEach((item: any) => {
                sub_section_temps.push({
                    subsection_template_id: item.id,
                    order_sub_index: item.order_sub_index,
                });
            });
        }
    });

    return [section_temps, sub_section_temps];
};

/* ------------------------------------------------- DOCX Export Function ------------------------------------------------- */
const exportToDocx = async (
    project: Project | null,
    bank: Banks | null,
    client: Clients[],
    checkedByGroup: Group[],
    images: string[]
) => {
    if (!project || !bank) {
        console.log('Project or bank data is missing');
        return;
    }

    try {
        const children: any[] = [];

        // Cover Page Section
        children.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Lenders Independent Engineers Report",
                        bold: true,
                        size: 32,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 300, after: 300, line: 300 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "a project of",
                        size: 24,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 200 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: project.name || "",
                        bold: true,
                        size: 36,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200, after: 400 }
            })
        );

        // Add bank info
        if (bank.name) {
            children.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: "For",
                            size: 24,
                        }),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 200 }
                }),
                new Paragraph({
                    children: [
                        new TextRun({
                            text: bank.name,
                            bold: true,
                            size: 32,
                        }),
                    ],
                    alignment: AlignmentType.CENTER,
                    spacing: { before: 100, after: 200 }
                })
            );
        }

        // Add company info
        children.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "By",
                        size: 24,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Adas Financial Research & Consulting Pvt. Ltd",
                        bold: true,
                        size: 28,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 100, after: 100 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Aditya Trade Center, Ameerpet, Hyderabad",
                        size: 22,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "www.atlasfin.in",
                        color: "0000FF",
                        underline: {},
                        size: 22,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 100 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: formatMonthYear(project.created_at),
                        size: 22,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 200, after: 300 }
            }),
            new PageBreak()
        );

        // Disclaimer Section
        children.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Disclaimer",
                        bold: true,
                        size: 32,
                        color: "032F5D",
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 300, after: 300 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "This Lenders' Independent Engineer (LIE) Report as of July 2025 has been prepared at the request from Canara Bank, Large Corporate Branch, Punjagutta, Hyderabad. The report contains proprietary and confidential information.",
                        size: 24,
                    }),
                ],
                spacing: { after: 200 }
            }),
            new PageBreak()
        );

        // About Section
        children.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "About Atlas Financial Research & Consulting (P) Ltd",
                        bold: true,
                        size: 28,
                        color: "032F5D",
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 300, after: 300 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Atlas Financial Research & Consulting (P) Ltd. is a distinguished company specializing in financial solutions and technical services essential for our client business's triumphant journey. With over a decade of experience, the Atlas team consists of esteemed professionals, including former senior executives from both public and private sector banks and seasoned technocrats. This wealth of expertise positions Atlas perfectly to cater to the diverse needs of the client.",
                        size: 24,
                    }),
                ],
                spacing: { after: 200 }
            }),
            new PageBreak()
        );

        // Table of Contents
        children.push(
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Table of Contents",
                        bold: true,
                        size: 32,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 300, after: 300 }
            })
        );

        // Add TOC items
        checkedByGroup.forEach((group, groupIndex) => {
            if (!group?.checked) return;

            children.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: group.title,
                            bold: true,
                            size: 22,
                        }),
                    ],
                    spacing: { before: 100, after: 50 }
                })
            );

            group.items.forEach((item, itemIndex) => {
                children.push(
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: `   ${item.label}`,
                                size: 20,
                            }),
                        ],
                        spacing: { before: 10, after: 10 }
                    })
                );
            });
        });

        children.push(new PageBreak());

        // Main Content Sections
        checkedByGroup.forEach((group, groupIndex) => {
            if (!group?.checked) return;

            children.push(
                new Paragraph({
                    children: [
                        new TextRun({
                            text: group.title,
                            bold: true,
                            size: 32,
                            color: "032F5D",
                        }),
                    ],
                    spacing: { before: 300, after: 200 }
                })
            );

            // Add section content
            if (Array.isArray(group.content)) {
                group.content.forEach((content: OlgptAny) => {
                    if (content.content_type === 'paragraph') {
                        children.push(
                            new Paragraph({
                                children: [
                                    new TextRun({
                                        text: content.content,
                                        size: 22,
                                    }),
                                ],
                                spacing: { before: 100, after: 100 }
                            })
                        );
                    }
                });
            }

            // Add subsections
            group.items.forEach((item, itemIndex) => {
                children.push(
                    new Paragraph({
                        children: [
                            new TextRun({
                                text: item.label,
                                bold: true,
                                size: 26,
                            }),
                        ],
                        spacing: { before: 200, after: 100 }
                    })
                );

                if (Array.isArray(item.content)) {
                    item.content.forEach((content: OlgptAny) => {
                        if (content.content_type === 'paragraph') {
                            children.push(
                                new Paragraph({
                                    children: [
                                        new TextRun({
                                            text: content.content,
                                            size: 22,
                                        }),
                                    ],
                                    spacing: { before: 50, after: 50 }
                                })
                            );
                        } else if (content.content_type === 'table_horizontal') {
                            // Create table
                            const table = new Table({
                                rows: [
                                    new TableRow({
                                        children: content.headers.map(header =>
                                            new TableCell({
                                                children: [new Paragraph({
                                                    children: [
                                                        new TextRun({
                                                            text: header,
                                                            bold: true,
                                                        })
                                                    ]
                                                })],
                                                shading: {
                                                    fill: "E8E8E8",
                                                },
                                            })
                                        ),
                                    }),
                                    ...content.rows.map(row =>
                                        new TableRow({
                                            children: row.map(cell =>
                                                new TableCell({
                                                    children: [new Paragraph({
                                                        children: [
                                                            new TextRun({
                                                                text: cell,
                                                            })
                                                        ]
                                                    })]
                                                })
                                            ),
                                        })
                                    ),
                                ],
                                width: {
                                    size: 100,
                                    type: WidthType.PERCENTAGE,
                                },
                            });

                            children.push(table);
                        }
                    });
                }
            });

            // Add page break between sections
            if (groupIndex < checkedByGroup.length - 1) {
                children.push(new PageBreak());
            }
        });

        // Annexures
        children.push(
            new PageBreak(),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "Annexures 2",
                        bold: true,
                        size: 32,
                        color: "2563EB",
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { before: 300, after: 200 }
            }),
            new Paragraph({
                children: [
                    new TextRun({
                        text: "just a subtitle",
                        size: 24,
                    }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: { after: 300 }
            })
        );

        // Create the document
        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        margin: {
                            top: 1440, // 1 inch
                            bottom: 1440,
                            left: 1440,
                            right: 1440
                        }
                    },
                },
                children: children,
            }],
        });

        // Generate and save the document
        const blob = await Packer.toBlob(doc);
        saveAs(blob, `${project.name?.replace(/[^a-z0-9]/gi, '_')}_LIE_Report_${new Date().toISOString().split('T')[0]}.docx`);

    } catch (error) {
        console.error('Error generating DOCX:', error);
        throw error;
    }
};

/* ------------------------------------------------- Page Components ------------------------------------------------- */
interface PageProps {
    id: string;
    children: React.ReactNode;
    pageNumber?: number;
    orientation?: 'portrait' | 'landscape';
    className?: string;
}

const Page: React.FC<PageProps> = memo(({ id, children, pageNumber, orientation = 'portrait', className = '' }) => {
    const pageRef = useRef<HTMLDivElement>(null);

    return (
        <Box
            ref={pageRef}
            id={id}
            className={`page ${className}`}
            sx={{
                backgroundColor: 'white',
                boxShadow: '0 0 20px rgba(0,0,0,0.1)',
                margin: orientation === 'landscape' ? '20px 40px' : '20px auto',
                padding: '30px',
                position: 'relative' as const,
                boxSizing: 'border-box' as const,
                breakInside: 'avoid' as const,
                overflow: 'hidden',
                width: orientation === 'landscape' ? '297mm' : '210mm',
                height: orientation === 'landscape' ? '210mm' : '297mm',
                minWidth: orientation === 'landscape' ? '297mm' : '210mm',
                minHeight: orientation === 'landscape' ? '210mm' : '297mm',
                '@media print': {
                    margin: '0',
                    padding: '25mm',
                    boxShadow: 'none',
                    pageBreakAfter: 'always' as const,
                    pageBreakInside: 'avoid' as const,
                    width: orientation === 'landscape' ? '297mm' : '210mm',
                    height: orientation === 'landscape' ? '210mm' : '297mm',
                }
            }}
        >
            {children}
            {pageNumber && (
                <Box
                    sx={{
                        position: 'absolute',
                        bottom: '15px',
                        right: '30px',
                        fontSize: '12px',
                        color: '#666',
                        fontFamily: 'Arial, sans-serif',
                        '@media print': {
                            position: 'fixed',
                            bottom: '15px',
                            right: '30px',
                        }
                    }}
                >
                    Page {pageNumber}
                </Box>
            )}
        </Box>
    );
});

interface ContentPageProps {
    content?: React.ReactNode;
    title?: string;
    pageId: string;
    pageNumber: number;
    orientation?: 'portrait' | 'landscape';
    isSubsection?: boolean;
    children?: React.ReactNode;
    sectionNumber?: string;
}

const ContentPage: React.FC<ContentPageProps> = memo(({
    content,
    children,
    title,
    pageId,
    pageNumber,
    orientation = 'portrait',
    isSubsection = false,
    sectionNumber = ""
}) => {
    return (
        <Page id={pageId} pageNumber={pageNumber} orientation={orientation}>
            <Box sx={{
                height: '100%',
                display: 'flex',
                flexDirection: 'column',
                position: 'relative'
            }}>
                {title && (
                    <Typography
                        variant={isSubsection ? "h5" : "h4"}
                        sx={{
                            fontWeight: 'bold',
                            color: '#032F5D',
                            mb: 3,
                            ...(isSubsection ? {
                                fontSize: '1.5rem'
                            } : {
                                fontSize: '1.75rem'
                            })
                        }}
                    >
                        {sectionNumber && !title.includes('continued') && `${sectionNumber} `}{title}
                    </Typography>
                )}
                <Box sx={{
                    flex: 1,
                    overflow: 'hidden'
                }}>
                    {children || content}
                </Box>
            </Box>
        </Page>
    );
});

/* ------------------------------------------------- UI Components ------------------------------------------------- */
const SectionHeader = memo(({ title }: { title: string }) => (
    <Box className="rounded-full" sx={{ color: 'common.black', px: 1, py: 1, mb: 1 }}>
        <Typography variant="subtitle2" className='font-extrabold text-[#032F5D]'>{title}</Typography>
    </Box>
));

interface SortableTaskItemProps {
    task: Task;
    group: Group;
    setCurrentCheck: (type: "section" | "subsection") => void;
    onToggle: (id: string) => void;
    allSection: SectionTemplates[];
    reportV: ReportVersions | null;
    safeFetch: (url: string) => Promise<any>;
    setAllSection: React.Dispatch<React.SetStateAction<SectionTemplates[]>>;
    setAllSubSection: React.Dispatch<React.SetStateAction<SubSectionTemplates[]>>;
    projectRS: ProjectReportSectionsP[];
    onGenerateContent: (sectionId: string, subsectionId: string, query?: string) => void;
    generatingMap: Record<string, boolean>;
    setSnackbar: React.Dispatch<React.SetStateAction<{
        open: boolean;
        message: string;
        severity: 'error' | 'warning' | 'success';
    }>>;
}

const SortableTaskItem = memo(({
    task,
    group,
    setCurrentCheck,
    onToggle,
    allSection,
    reportV,
    safeFetch,
    setAllSection,
    setAllSubSection,
    projectRS,
    onGenerateContent,
    generatingMap,
    setSnackbar,
}: SortableTaskItemProps) => {
    const { attributes, listeners, setNodeRef, setActivatorNodeRef, transform, transition } =
        useSortable({ id: task.id, data: { type: 'item' } });

    const style = { transform: CSS.Transform.toString(transform), transition };
    const [generating, setGenerating] = useState(false);
    const [showQueryDialog, setShowQueryDialog] = useState(false);
    const [pendingGeneration, setPendingGeneration] = useState<{ sectionId: string, subsectionId: string } | null>(null);

    const handleDelete = useCallback(async () => {
        await fetch(`${config.backendURL}/api/subsection-templates/${task.id}`, {
            method: "DELETE",
            credentials: "include"
        });

        const prs = projectRS.find(p => p.section_template_id === task.id);
        if (prs) {
            await fetch(`${config.backendURL}/api/project-report-sections/${prs.id}`, {
                method: "PUT",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ ...prs, is_deleted: true })
            });
        }

        setAllSection(allSection.filter(s => s.id !== task.id));
    }, [task.id, projectRS, allSection, setAllSection]);

    const onChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
        e.stopPropagation();
        setCurrentCheck("subsection");
        onToggle(task.id);
    }, [onToggle, task.id, setCurrentCheck]);

    const handleGenerateClick = useCallback(() => {
        setPendingGeneration({ sectionId: group.id, subsectionId: task.id });
        setShowQueryDialog(true);
    }, [group.id, task.id]);

    const handleQuerySubmit = useCallback(async (query: string) => {
        if (!pendingGeneration) return;

        setGenerating(true);
        setShowQueryDialog(false);
        try {
            await onGenerateContent(pendingGeneration.sectionId, pendingGeneration.subsectionId, query);
        } finally {
            setGenerating(false);
            setPendingGeneration(null);
        }
    }, [pendingGeneration, onGenerateContent]);

    return (
        <>
            <ListItem ref={setNodeRef} style={style} className="rounded-lg" sx={{ px: 1 }} {...attributes}>
                <ListItemIcon sx={{ minWidth: 36 }}>
                    <Checkbox
                        edge="start"
                        checked={!!task.checked}
                        onChange={onChange}
                        tabIndex={-1}
                        disableRipple
                    />
                </ListItemIcon>
                <ListItemText primary={task.label} secondary={task.note} />

                <Tooltip title="Generate Content">
                    <IconButton
                        size="small"
                        onClick={handleGenerateClick}
                        // disabled={
                        //     generating ||
                        //     !task.checked ||
                        //     !!generatingMap[task.main_id ?? '']
                        // }
                        disabled={generating || !group.checked}
                        sx={{ color: '#032F5D' }}
                    >
                        {generating ? (
                            <CircularProgress size={20} />
                        ) : (
                            <AutoAwesome fontSize="small" />
                        )}
                    </IconButton>
                </Tooltip>

                <button className='text-red-500' onClick={handleDelete}>
                    <Delete />
                </button>
                <IconButton
                    aria-label="Drag task"
                    size="small"
                    sx={{ cursor: 'grab' }}
                    ref={setActivatorNodeRef}
                    {...listeners}
                >
                    <DragIndicatorRounded fontSize="small" />
                </IconButton>
            </ListItem>

            <GenerationQueryDialog
                open={showQueryDialog}
                onClose={() => {
                    setShowQueryDialog(false);
                    setPendingGeneration(null);
                }}
                onSubmit={handleQuerySubmit}
                title={`Generate Content for: ${task.label}`}
                loading={generating}
            />
        </>
    );
});

interface SortableGroupProps {
    group: Group;
    setCurrentCheck: (type: "section" | "subsection") => void;
    onToggle: (id: string) => void;
    onToggleCheck: (id: string) => void;
    onToggleTask: (id: string) => void;
    allSection: SectionTemplates[];
    reportV: ReportVersions | null;
    safeFetch: (url: string) => Promise<any>;
    setAllSection: React.Dispatch<React.SetStateAction<SectionTemplates[]>>;
    setAllSubSection: React.Dispatch<React.SetStateAction<SubSectionTemplates[]>>;
    projectRS: ProjectReportSectionsP[];
    project: Project | null;
    onGenerateContent: (sectionId: string, subsectionId?: string, query?: string) => void;
    generatingMap: Record<string, boolean>;
    setSnackbar: React.Dispatch<React.SetStateAction<{
        open: boolean;
        message: string;
        severity: 'error' | 'warning' | 'success';
    }>>;
}

const SortableGroup = memo(({
    group,
    setCurrentCheck,
    onToggle,
    onToggleCheck,
    onToggleTask,
    allSection,
    reportV,
    safeFetch,
    setAllSection,
    setAllSubSection,
    projectRS,
    project,
    onGenerateContent,
    generatingMap,
    setSnackbar,
}: SortableGroupProps) => {
    const [newSubName, setNewSubName] = useState<Record<string, string>>({});
    const [generating, setGenerating] = useState(false);
    const [showQueryDialog, setShowQueryDialog] = useState(false);
    const [pendingGeneration, setPendingGeneration] = useState<string | null>(null);

    const { attributes, listeners, setNodeRef, setActivatorNodeRef, transform, transition } =
        useSortable({ id: group.id, data: { type: 'container' } });

    const style = { transform: CSS.Transform.toString(transform), transition };

    const handleDelete = useCallback(async () => {
        await fetch(`${config.backendURL}/api/section-templates/${group.id}`, {
            method: "DELETE",
            credentials: "include"
        });

        const prs = projectRS.find(p => p.section_template_id === group.id);
        if (prs) {
            await fetch(`${config.backendURL}/api/project-report-sections/${prs.id}`, {
                method: "PUT",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ ...prs, is_deleted: true })
            });
        }

        setAllSection(allSection.filter(s => s.id !== group.id));
    }, [group.id, projectRS, allSection, setAllSection]);

    const handleAddSubsection = useCallback(async (groupId: string) => {
        const name = newSubName[groupId]?.trim();
        if (!name) return;
        if (!project || !reportV) return;

        setNewSubName(prev => ({ ...prev, [groupId]: "" }));

        const newSub = await fetch(`${config.backendURL}/api/subsection-templates`, {
            method: "POST",
            credentials: "include",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                section_template_id: groupId,
                name,
                default_content: JSON.stringify([{
                    "content": "The project funding is well-balanced. The next debt tranche (INR 30 Cr) is tied to the physical completion of the 20th slab on all four towers and 75% MEP completion, expected in February 2026. The customer collection trajectory is healthy and exceeds the equity requirement.",
                    "content_type": "paragraph"
                }]),
                is_mandatory: false,
                project_type: project.project_type,
                order_suggestion: group.items.length + 1,
            }),
        }).then(r => r.json());

        await fetch(`${config.backendURL}/api/project-report-sections/from-template`, {
            method: "POST",
            credentials: "include",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                report_version_id: reportV ? reportV?.id : null,
                section_template_id: groupId,
                subsection_template_id: newSub.id,
            }),
        });

        const updatedSubs = await safeFetch(`/api/subsection-templates`);
        setAllSubSection(updatedSubs);
    }, [newSubName, group.items.length, reportV, safeFetch, setAllSubSection]);

    const onCheckChange = useCallback((e: React.ChangeEvent<HTMLInputElement>) => {
        e.stopPropagation();
        setCurrentCheck("section");
        onToggleCheck(group.id);
    }, [onToggleCheck, group.id, setCurrentCheck]);

    const handleGenerateClick = useCallback(() => {
        setPendingGeneration(group.id);
        setShowQueryDialog(true);
    }, [group.id]);

    const handleQuerySubmit = useCallback(async (query: string) => {
        if (!pendingGeneration) return;

        setGenerating(true);
        setShowQueryDialog(false);
        try {
            await onGenerateContent(pendingGeneration, undefined, query);
        } finally {
            setGenerating(false);
            setPendingGeneration(null);
        }
    }, [pendingGeneration, onGenerateContent]);

    if (group.title === "Cover Page") return null;

    return (
        <>
            <Box ref={setNodeRef} style={style} className="mb-1" {...attributes}>
                <Stack direction="row" alignItems="center" justifyContent="start" sx={{ mb: 1 }} className='pe-4 bg-primary-800/5'>
                    <Stack direction="row" alignItems="center" spacing={0.5} sx={{ flexGrow: 1 }}>
                        <IconButton
                            ref={setActivatorNodeRef}
                            {...listeners}
                            size="small"
                            aria-label="Drag group"
                            sx={{ cursor: 'grab' }}
                        >
                            <DragIndicatorRounded fontSize="small" />
                        </IconButton>
                        <Checkbox
                            edge="start"
                            checked={!!group.checked}
                            onChange={onCheckChange}
                            tabIndex={-1}
                            disableRipple
                        />
                        <SectionHeader title={group.title} />

                        <Tooltip title="Generate Section Content">
                            <IconButton
                                size="small"
                                onClick={handleGenerateClick}
                                // disabled={
                                //     generating ||
                                //     !group.checked ||
                                //     !!generatingMap[group.main_id ?? '']
                                // }
                                disabled={generating || !group.checked}
                                sx={{ color: '#032F5D' }}
                            >
                                {generating ? (
                                    <CircularProgress size={20} />
                                ) : (
                                    <AutoAwesome fontSize="small" />
                                )}
                            </IconButton>
                        </Tooltip>
                    </Stack>
                    <button className='text-red-500' onClick={handleDelete}>
                        <Delete />
                    </button>
                    <IconButton
                        aria-label={group.expanded ? 'Collapse group' : 'Expand group'}
                        onClick={() => onToggle(group.id)}
                        edge="end"
                        sx={{
                            transform: group.expanded ? 'rotate(180deg)' : 'rotate(0deg)',
                            transition: 'transform 150ms'
                        }}
                    >
                        <ExpandMore />
                    </IconButton>
                </Stack>
                <Collapse in={!!group.expanded} unmountOnExit className="ps-3">
                    {group.items.length > 0 && (
                        <SortableContext items={group.items.map(t => t.id)} strategy={verticalListSortingStrategy}>
                            <List dense className="space-y-2">
                                {group.items.map(t => (
                                    <SortableTaskItem
                                        key={t.id}
                                        task={t}
                                        group={group}
                                        setCurrentCheck={setCurrentCheck}
                                        allSection={allSection}
                                        reportV={reportV}
                                        onToggle={onToggleTask}
                                        safeFetch={safeFetch}
                                        setAllSection={setAllSection}
                                        setAllSubSection={setAllSubSection}
                                        projectRS={projectRS}
                                        onGenerateContent={onGenerateContent}
                                        generatingMap={generatingMap}
                                        setSnackbar={setSnackbar}
                                    />
                                ))}
                            </List>
                        </SortableContext>
                    )}
                    <div className="mt-2 pl-4 grid grid-cols-[1fr_40px] gap-2 items-center">
                        <input
                            type="text"
                            placeholder="Enter subsection name"
                            className="w-full border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                            value={newSubName[group.id] || ""}
                            onChange={(e) => setNewSubName(prev => ({ ...prev, [group.id]: e.target.value }))}
                        />
                        <button
                            className="w-full h-full py-1 text-xs bg-green-600 text-white rounded hover:bg-green-700 disabled:bg-gray-400"
                            disabled={!newSubName[group.id]?.trim()}
                            onClick={() => handleAddSubsection(group.id)}
                        >
                            +
                        </button>
                    </div>
                </Collapse>
            </Box>

            <GenerationQueryDialog
                open={showQueryDialog}
                onClose={() => {
                    setShowQueryDialog(false);
                    setPendingGeneration(null);
                }}
                onSubmit={handleQuerySubmit}
                title={`Generate Content for: ${group.title}`}
                loading={generating}
            />
        </>
    );
});

/* ------------------------------------------------- Table of Contents ------------------------------------------------- */
interface TableOfContentsProps {
    checkedByGroup: Group[];
    pageNumbers: Record<string, { start: number; end?: number }>;
    scrollToPage: (pageId: string) => void;
}

const TableOfContents: React.FC<TableOfContentsProps> = memo(({ checkedByGroup, pageNumbers, scrollToPage }) => {
    const [hoveredItem, setHoveredItem] = useState<string | null>(null);

    return (
        <Box sx={{ display: 'flex', flexDirection: 'column', height: '100%', width: '100%' }}>
            <Typography variant="h4" gutterBottom className='mb-4 text-center font-bold' sx={{ color: '#032F5D' }}>
                Table of Contents
            </Typography>

            <Box sx={{ flex: 1, overflow: 'hidden', width: '100%' }}>
                <Box sx={{
                    height: '100%',
                    overflow: 'auto',
                    pr: 1,
                    width: '100%'
                }}>
                    {checkedByGroup.length === 0 && (
                        <Typography variant="body2" color="text.secondary" textAlign="center">
                            No items selected yet. Check tasks in the sidebar to see content here.
                        </Typography>
                    )}

                    {checkedByGroup.map((group, groupIndex) => {
                        if (!group?.checked) return null;

                        const groupPageId = `section-${group.id}-page-1`;
                        const pageInfo = pageNumbers[groupPageId];

                        return (
                            <Box key={group.id} sx={{ mb: 3, width: '100%' }}>
                                <div className='flex justify-between items-center gap-3 mb-2 w-full'>
                                    <button
                                        onClick={() => scrollToPage(groupPageId)}
                                        onMouseEnter={() => setHoveredItem(groupPageId)}
                                        onMouseLeave={() => setHoveredItem(null)}
                                        className={`text-left w-[max-content] flex-shrink-0 mb-0 transition-all duration-200 ${hoveredItem === groupPageId
                                            ? 'text-blue-600 underline decoration-2 decoration-blue-500'
                                            : 'text-gray-900'
                                            }`}
                                        style={{ cursor: 'pointer' }}
                                    >
                                        <Typography variant="h6" className='font-bold' sx={{ fontSize: '1.1rem' }}>
                                            {groupIndex + 1}. {group.title}
                                        </Typography>
                                    </button>
                                    <div className='flex-grow h-[2px] bg-gray-200'></div>
                                    <button
                                        onClick={() => scrollToPage(groupPageId)}
                                        onMouseEnter={() => setHoveredItem(groupPageId)}
                                        onMouseLeave={() => setHoveredItem(null)}
                                        className={`page-no text-[0.875rem] w-[max-content] flex-shrink-0 transition-all duration-200 ${hoveredItem === groupPageId
                                            ? 'text-blue-600 font-bold'
                                            : 'text-gray-700'
                                            }`}
                                        style={{ cursor: 'pointer', minWidth: '40px', textAlign: 'right' }}
                                    >
                                        {pageInfo ? pageInfo.start : 1}
                                        {pageInfo?.end && pageInfo.end > pageInfo.start ? `-${pageInfo.end}` : ''}
                                    </button>
                                </div>
                                {group.items.map((item, itemIndex) => {
                                    const itemPageId = `subsection-${item.id}-page-1`;
                                    const itemPageInfo = pageNumbers[itemPageId];

                                    return (
                                        <Box key={item.id} sx={{ mb: 1.5 }} className='flex justify-between items-center gap-3 ps-4 w-full'>
                                            <button
                                                onClick={() => scrollToPage(itemPageId)}
                                                onMouseEnter={() => setHoveredItem(itemPageId)}
                                                onMouseLeave={() => setHoveredItem(null)}
                                                className={`text-left w-[max-content] flex-shrink-0 mb-0 transition-all duration-200 ${hoveredItem === itemPageId
                                                    ? 'text-blue-600 underline decoration-1 decoration-blue-400'
                                                    : 'text-gray-800'
                                                    }`}
                                                style={{ cursor: 'pointer', fontSize: '0.875rem' }}
                                            >
                                                <Typography variant="subtitle1" sx={{ fontSize: '0.875rem' }}>
                                                    {groupIndex + 1}.{itemIndex + 1} {item.label}
                                                </Typography>
                                            </button>
                                            <div className='flex-grow h-[1px] bg-gray-100'></div>
                                            <button
                                                onClick={() => scrollToPage(itemPageId)}
                                                onMouseEnter={() => setHoveredItem(itemPageId)}
                                                onMouseLeave={() => setHoveredItem(null)}
                                                className={`page-no text-[0.875rem] w-[max-content] flex-shrink-0 transition-all duration-200 ${hoveredItem === itemPageId
                                                    ? 'text-blue-600 font-bold'
                                                    : 'text-gray-600'
                                                    }`}
                                                style={{ cursor: 'pointer', minWidth: '40px', textAlign: 'right' }}
                                            >
                                                {itemPageInfo ? itemPageInfo.start : 1}
                                                {itemPageInfo?.end && itemPageInfo.end > itemPageInfo.start ? `-${itemPageInfo.end}` : ''}
                                            </button>
                                        </Box>
                                    );
                                })}
                            </Box>
                        );
                    })}
                </Box>
            </Box>
        </Box>
    );
});

/* ------------------------------------------------- Content Splitter ------------------------------------------------- */
const splitContentIntoPages = (content: any[], pageIdPrefix: string): { pages: any[][]; pageIds: string[] } => {
    if (!Array.isArray(content) || content.length === 0) {
        return { pages: [[]], pageIds: [`${pageIdPrefix}-page-1`] };
    }

    const pages: any[][] = [];
    const pageIds: string[] = [];
    let currentPage: any[] = [];
    let currentPageNumber = 1;

    const estimateContentSize = (item: any): number => {
        if (item.content_type === 'paragraph') {
            return item.content ? item.content.length * 0.5 : 100;
        } else if (item.content_type === 'table_horizontal') {
            return 1500 + (item.rows?.length || 0) * 100;
        } else if (item.content_type === 'table_vertical') {
            return 1000 + (item.content?.length || 0) * 80;
        }
        return 100;
    };

    let currentPageSize = 0;
    const maxPageSize = 2800;

    content.forEach((item, index) => {
        const itemSize = estimateContentSize(item);

        if (currentPageSize + itemSize > maxPageSize && currentPage.length > 0) {
            pages.push([...currentPage]);
            pageIds.push(`${pageIdPrefix}-page-${currentPageNumber}`);
            currentPage = [item];
            currentPageSize = itemSize;
            currentPageNumber++;
        } else {
            currentPage.push(item);
            currentPageSize += itemSize;
        }
    });

    if (currentPage.length > 0) {
        pages.push(currentPage);
        pageIds.push(`${pageIdPrefix}-page-${currentPageNumber}`);
    }

    return { pages, pageIds };
};

/* ------------------------------------------------- TOC Splitter - Enterprise Level ------------------------------------------------- */
const splitTOCIntoPages = (
    checkedByGroup: Group[]
): {
    pages: {
        items: {
            group: Group;
            item?: Task;
            sectionNumber: string;
            subsectionNumber?: string;
        }[];
        pageNumber: number;
    }[];
    totalPages: number;
    sectionNumbering: Record<string, string>;
    subsectionNumbering: Record<string, string>;
} => {
    if (checkedByGroup.length === 0) {
        return { pages: [], totalPages: 0, sectionNumbering: {}, subsectionNumbering: {} };
    }

    const pages: {
        items: {
            group: Group;
            item?: Task;
            sectionNumber: string;
            subsectionNumber?: string;
        }[];
        pageNumber: number;
    }[] = [];

    const sectionNumbering: Record<string, string> = {};
    const subsectionNumbering: Record<string, string> = {};

    checkedByGroup.forEach((group, groupIndex) => {
        if (!group?.checked) return;
        sectionNumbering[group.id] = (groupIndex + 1).toString();

        group.items.forEach((item, itemIndex) => {
            if (item.checked) {
                subsectionNumbering[item.id] = `${groupIndex + 1}.${itemIndex + 1}`;
            }
        });
    });

    const ITEMS_PER_PAGE = 25;
    const GROUP_ITEM_WEIGHT = 1.5;

    const allTOCItems: {
        group: Group;
        item?: Task;
        weight: number;
        sectionNumber: string;
        subsectionNumber?: string;
    }[] = [];

    checkedByGroup.forEach((group, groupIndex) => {
        if (!group?.checked) return;

        const sectionNum = sectionNumbering[group.id];

        allTOCItems.push({
            group,
            item: undefined,
            weight: GROUP_ITEM_WEIGHT,
            sectionNumber: sectionNum
        });

        group.items.forEach((item, itemIndex) => {
            if (item.checked) {
                allTOCItems.push({
                    group,
                    item,
                    weight: 1,
                    sectionNumber: sectionNum,
                    subsectionNumber: subsectionNumbering[item.id]
                });
            }
        });
    });

    let currentPageItems: {
        group: Group;
        item?: Task;
        sectionNumber: string;
        subsectionNumber?: string;
    }[] = [];
    let currentPageWeight = 0;
    let pageCount = 0;

    for (const tocItem of allTOCItems) {
        if (currentPageWeight + tocItem.weight > ITEMS_PER_PAGE && currentPageItems.length > 0) {
            pages.push({
                items: [...currentPageItems],
                pageNumber: pages.length + 1
            });

            currentPageItems = [tocItem];
            currentPageWeight = tocItem.weight;
            pageCount++;
        } else {
            currentPageItems.push(tocItem);
            currentPageWeight += tocItem.weight;
        }
    }

    if (currentPageItems.length > 0) {
        pages.push({
            items: currentPageItems,
            pageNumber: pages.length + 1
        });
        pageCount++;
    }

    return {
        pages,
        totalPages: pageCount,
        sectionNumbering,
        subsectionNumbering
    };
};

interface TOCPageProps {
    items: {
        group: Group;
        item?: Task;
        sectionNumber: string;
        subsectionNumber?: string;
    }[];
    pageNumbers: Record<string, { start: number; end?: number }>;
    scrollToPage: (pageId: string) => void;
    pageNumber: number;
    totalPages: number;
}

const TOCPage: React.FC<TOCPageProps> = memo(({
    items,
    pageNumbers,
    scrollToPage,
    pageNumber,
    totalPages
}) => {
    const [hoveredItem, setHoveredItem] = useState<string | null>(null);

    return (
        <Box sx={{
            display: 'flex',
            flexDirection: 'column',
            height: '100%',
            width: '100%',
            position: 'relative',
            zIndex: 1000
        }}>
            <Typography
                variant="h4"
                gutterBottom
                className='mb-4 text-center font-bold'
                sx={{ color: '#032F5D' }}
            >
                Table of Contents {totalPages > 1 ? `(Page ${pageNumber} of ${totalPages})` : ''}
            </Typography>

            <Box sx={{ flex: 1, overflow: 'hidden', width: '100%' }}>
                <Box sx={{
                    height: '100%',
                    overflow: 'auto',
                    pr: 1,
                    width: '100%'
                }}>
                    {items.map((tocItem, index) => {
                        if (!tocItem.item) {
                            const groupPageId = `section-${tocItem.group.id}-page-1`;
                            const pageInfo = pageNumbers[groupPageId];

                            return (
                                <Box key={`group-${tocItem.group.id}`} sx={{ mb: 2, width: '100%' }}>
                                    <div className='flex justify-between items-center gap-3 mb-2 w-full'>
                                        <button
                                            onClick={() => scrollToPage(groupPageId)}
                                            onMouseEnter={() => setHoveredItem(groupPageId)}
                                            onMouseLeave={() => setHoveredItem(null)}
                                            className={`text-left w-[max-content] flex-shrink-0 mb-0 transition-all duration-200 ${hoveredItem === groupPageId
                                                ? 'text-blue-600 underline decoration-2 decoration-blue-500'
                                                : 'text-gray-900'
                                                }`}
                                            style={{ cursor: 'pointer' }}
                                        >
                                            <Typography variant="h6" className='font-bold' sx={{ fontSize: '1.1rem' }}>
                                                {tocItem.sectionNumber}. {tocItem.group.title}
                                            </Typography>
                                        </button>
                                        <div className='flex-grow h-[2px] bg-gray-200'></div>
                                        <button
                                            onClick={() => scrollToPage(groupPageId)}
                                            onMouseEnter={() => setHoveredItem(groupPageId)}
                                            onMouseLeave={() => setHoveredItem(null)}
                                            className={`page-no text-[0.875rem] w-[max-content] flex-shrink-0 transition-all duration-200 ${hoveredItem === groupPageId
                                                ? 'text-blue-600 font-bold'
                                                : 'text-gray-700'
                                                }`}
                                            style={{ cursor: 'pointer', minWidth: '40px', textAlign: 'right' }}
                                        >
                                            {pageInfo ? pageInfo.start : 1}
                                            {pageInfo?.end && pageInfo.end > pageInfo.start ? `-${pageInfo.end}` : ''}
                                        </button>
                                    </div>
                                </Box>
                            );
                        } else {
                            const itemPageId = `subsection-${tocItem.item.id}-page-1`;
                            const itemPageInfo = pageNumbers[itemPageId];

                            return (
                                <Box key={`item-${tocItem.item.id}`} sx={{ mb: 1.5 }} className='flex justify-between items-center gap-3 ps-4 w-full'>
                                    <button
                                        onClick={() => scrollToPage(itemPageId)}
                                        onMouseEnter={() => setHoveredItem(itemPageId)}
                                        onMouseLeave={() => setHoveredItem(null)}
                                        className={`text-left w-[max-content] flex-shrink-0 mb-0 transition-all duration-200 ${hoveredItem === itemPageId
                                            ? 'text-blue-600 underline decoration-1 decoration-blue-400'
                                            : 'text-gray-800'
                                            }`}
                                        style={{ cursor: 'pointer', fontSize: '0.875rem' }}
                                    >
                                        <Typography variant="subtitle1" sx={{ fontSize: '0.875rem' }}>
                                            {tocItem.subsectionNumber} {tocItem.item.label}
                                        </Typography>
                                    </button>
                                    <div className='flex-grow h-[1px] bg-gray-100'></div>
                                    <button
                                        onClick={() => scrollToPage(itemPageId)}
                                        onMouseEnter={() => setHoveredItem(itemPageId)}
                                        onMouseLeave={() => setHoveredItem(null)}
                                        className={`page-no text-[0.875rem] w-[max-content] flex-shrink-0 transition-all duration-200 ${hoveredItem === itemPageId
                                            ? 'text-blue-600 font-bold'
                                            : 'text-gray-600'
                                            }`}
                                        style={{ cursor: 'pointer', minWidth: '40px', textAlign: 'right' }}
                                    >
                                        {itemPageInfo ? itemPageInfo.start : 1}
                                        {itemPageInfo?.end && itemPageInfo.end > itemPageInfo.start ? `-${itemPageInfo.end}` : ''}
                                    </button>
                                </Box>
                            );
                        }
                    })}
                </Box>
            </Box>
        </Box>
    );
});

/* ------------------------------------------------- Grammar Check Function ------------------------------------------------- */
const grammarCheck = async (text: string): Promise<string> => {
    try {
        const response = await fetch("http://13.127.94.207:8000/api/grammar-check", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
                task: "grammar_check",
                text: text,
            })
        });

        if (!response.ok) {
            throw new Error("Grammar check failed");
        }

        const result = await response.json();
        return result.data.corrected_text || "";
    } catch (error) {
        console.error("Grammar check error:", error);
        throw error;
    }
};

// Update the ReportViewerProps interface
interface ReportViewerProps {
    project: Project | null;
    bank: Banks | null;
    client: Clients[];
    checkedByGroup: Group[];
    renderOlResult: (
        key: string,
        content: OlgptAny,
        sectionId: string,
        index: number,
        fullContent: OlgptAny[],
        onOpenCommentModal?: (entityType: string, entityId: string) => void
    ) => React.ReactNode;
    onExportDocx: () => Promise<void>;
    onExportPdf: () => void;
}

const ReportViewer: React.FC<ReportViewerProps> = memo(({
    project,
    bank,
    client,
    checkedByGroup,
    renderOlResult,
    onExportDocx,
    onExportPdf
}) => {
    const [viewMode, setViewMode] = useState<'preview' | 'print'>('preview');
    const [pageNumbers, setPageNumbers] = useState<Record<string, { start: number; end?: number }>>({});
    const [exporting, setExporting] = useState(false);
    const [exportFormat, setExportFormat] = useState<'docx' | 'pdf' | null>(null);
    const [zoomLevel, setZoomLevel] = useState<number>(0.8);
    const reportPagesRef = useRef<HTMLDivElement>(null);
    const [projectPhotos, setProjectPhotos] = useState<any[]>([]);
    const [showControls, setShowControls] = useState(true);
    const [isPending, startTransition] = useTransition();
    const [openComments, setOpenComments] = useState(false);
    const [selectedEntity, setSelectedEntity] = useState({ type: "", id: "" });
    const [editDialogOpen, setEditDialogOpen] = useState(false);
    const [editValue, setEditValue] = useState("");
    const [editData, setEditData] = useState({ value: [] as any[], id: "", index: -1 });
    const [loading, setLoading] = useState(false);
    const safeFetch = useSafeFetch();

    useEffect(() => {
        if (!project?.id) return;

        const loadProjectPhotos = async () => {
            try {
                const data = await safeFetch(
                    `/api/photos?project_id=${project.id}`
                );

                setProjectPhotos(Array.isArray(data) ? data : []);
            } catch (e) {
                console.error("Failed to load project photos", e);
                setProjectPhotos([]);
            }
        };

        loadProjectPhotos();
    }, [project?.id, safeFetch]);

    const annexurePhotos = useMemo(() => {
        return projectPhotos
            .filter(p => p.file_url)
            .sort(
                (a, b) =>
                    new Date(b.created_at).getTime() -
                    new Date(a.created_at).getTime()
            )
            .slice(0, 8); // LIMIT → fast UI
    }, [projectPhotos]);

    const zoomOptions = [
        { label: '50%', value: 0.5 },
        { label: '75%', value: 0.75 },
        { label: '80%', value: 0.8 },
        { label: '90%', value: 0.9 },
        { label: '100%', value: 1.0 },
        { label: '125%', value: 1.25 },
        { label: '150%', value: 1.5 },
    ];

    const debouncedSetPageNumbers = useRef<NodeJS.Timeout | null>(null);
    const {
        pages: tocPages,
        totalPages: tocTotalPages,
        sectionNumbering,
        subsectionNumbering
    } = useMemo(() =>
        splitTOCIntoPages(checkedByGroup),
        [checkedByGroup]
    );

    const handleOpenCommentModal = useCallback((entityType: string, entityId: string) => {
        setSelectedEntity({ type: entityType, id: entityId });
        setOpenComments(true);
    }, []);

    useEffect(() => {
        if (debouncedSetPageNumbers.current) {
            clearTimeout(debouncedSetPageNumbers.current);
        }

        debouncedSetPageNumbers.current = setTimeout(() => {
            startTransition(() => {
                const newPageNumbers: Record<string, { start: number; end?: number }> = {};
                let currentPage = 1;

                newPageNumbers['cover-page'] = { start: currentPage };
                currentPage++;

                newPageNumbers['disclaimer-page'] = { start: currentPage };
                currentPage++;

                newPageNumbers['about-page'] = { start: currentPage };
                currentPage++;

                tocPages.forEach((tocPage, index) => {
                    const pageId = `toc-page-${tocPage.pageNumber}`;
                    newPageNumbers[pageId] = { start: currentPage };
                    currentPage++;
                });

                checkedByGroup.forEach((group, groupIndex) => {
                    if (!group?.checked) return;

                    const sectionPages = splitContentIntoPages(group.content || [], `section-${group.id}`);
                    const groupPageId = `section-${group.id}-page-1`;
                    newPageNumbers[groupPageId] = {
                        start: currentPage,
                        end: currentPage + sectionPages.pages.length - 1
                    };
                    currentPage += sectionPages.pages.length;

                    group.items.forEach((item, itemIndex) => {
                        if (!item.checked) return;

                        const itemPages = splitContentIntoPages(item.content || [], `subsection-${item.id}`);
                        const itemPageId = `subsection-${item.id}-page-1`;
                        newPageNumbers[itemPageId] = {
                            start: currentPage,
                            end: currentPage + itemPages.pages.length - 1
                        };
                        currentPage += itemPages.pages.length;
                    });
                });

                newPageNumbers['annexures-page'] = { start: currentPage };

                setPageNumbers(newPageNumbers);
            });
        }, 100);

        return () => {
            if (debouncedSetPageNumbers.current) {
                clearTimeout(debouncedSetPageNumbers.current);
            }
        };
    }, [checkedByGroup, tocPages, startTransition]);

    const scrollToPage = useCallback((pageId: string) => {
        const element = document.getElementById(pageId);
        if (element) {
            element.scrollIntoView({
                behavior: 'smooth',
                block: 'start'
            });
        }
    }, []);

    const handleExport = useCallback(async (format: 'docx' | 'pdf') => {
        setExportFormat(format);
        setExporting(true);
        try {
            if (format === 'docx') {
                await onExportDocx();
            } else {
                onExportPdf();
            }
        } catch (error) {
            console.error('Export failed:', error);
            console.log(`Export to ${format.toUpperCase()} failed. Please try again.`);
        } finally {
            setExporting(false);
            setExportFormat(null);
        }
    }, [onExportDocx, onExportPdf]);

    const handleZoomIn = useCallback(() => {
        setZoomLevel(prev => Math.min(prev + 0.1, 2.0));
    }, []);

    const handleZoomOut = useCallback(() => {
        setZoomLevel(prev => Math.max(prev - 0.1, 0.3));
    }, []);

    const handleZoomReset = useCallback(() => {
        setZoomLevel(0.8);
    }, []);

    const toggleControls = useCallback(() => {
        setShowControls(!showControls);
    }, [showControls]);

    useEffect(() => {
        const handleKeyDown = (e: KeyboardEvent) => {
            if ((e.ctrlKey || e.metaKey) && e.key === '0') {
                e.preventDefault();
                handleZoomReset();
            }
            if ((e.ctrlKey || e.metaKey) && e.key === 'h') {
                e.preventDefault();
                toggleControls();
            }
        };

        const handleWheel = (e: WheelEvent) => {
            if (e.ctrlKey || e.metaKey) {
                e.preventDefault();
                if (e.deltaY < 0) {
                    handleZoomIn();
                } else {
                    handleZoomOut();
                }
            }
        };

        window.addEventListener('keydown', handleKeyDown);
        window.addEventListener('wheel', handleWheel, { passive: false });

        return () => {
            window.removeEventListener('keydown', handleKeyDown);
            window.removeEventListener('wheel', handleWheel);
        };
    }, [handleZoomIn, handleZoomOut, handleZoomReset, toggleControls]);

    const handleGrammarUpdate = async () => {
        if (!editData.id || editData.index === -1) return;

        const updatedContent = [...editData.value];
        updatedContent[editData.index] = {
            ...updatedContent[editData.index],
            content: editValue
        };

        try {
            const res = await fetch(`${config.backendURL}/api/project-report-sections/${editData.id}`, {
                method: "PUT",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({ content: updatedContent }),
            });

            if (res.ok) {
                // Refresh data if needed
            }
        } catch (e) {
            console.error("Failed to update content:", e);
        } finally {
            setEditDialogOpen(false);
        }
    };

    return (
        <Box sx={{
            width: '100%',
            maxHeight: '100vh',
            backgroundColor: '#f5f5f5',
            overflow: 'scroll',
            display: 'flex',
            flexDirection: 'column'
        }}>
            <CommentModal
                open={openComments}
                onClose={() => setOpenComments(false)}
                entityType={selectedEntity.type}
                entityId={selectedEntity.id}
                safeFetch={safeFetch}
                selectedEntity={selectedEntity}
                grammarCheck={grammarCheck}
                loading={loading}
                setLoading={setLoading}
            />

            <Dialog open={editDialogOpen} onClose={() => setEditDialogOpen(false)}>
                <DialogTitle>Edit Summary</DialogTitle>
                <DialogContent className='min-w-[500px]'>
                    <TextField
                        autoFocus
                        margin="dense"
                        label="Enter new summary"
                        type="text"
                        fullWidth
                        variant="outlined"
                        multiline
                        minRows={6}
                        value={editValue}
                        onChange={(e) => setEditValue(e.target.value)}
                    />
                </DialogContent>
                <DialogActions>
                    <Button onClick={() => setEditDialogOpen(false)} color="inherit">
                        Cancel
                    </Button>
                    <Button onClick={handleGrammarUpdate} variant="contained" color="primary">
                        Update
                    </Button>
                </DialogActions>
            </Dialog>

            <Box sx={{
                position: 'sticky',
                top: 0,
                zIndex: 5,
                backgroundColor: 'white',
                padding: 1,
                marginBottom: 1,
                boxShadow: 1,
                display: 'flex',
                justifyContent: 'space-between',
                alignItems: 'center',
                flexShrink: 0
            }}>
                <Box sx={{ display: 'flex', gap: 1, alignItems: 'center' }}>
                    <Typography variant="subtitle2">Report Viewer</Typography>
                    <Chip
                        label={viewMode === 'preview' ? "Preview" : "Print"}
                        color={viewMode === 'preview' ? "primary" : "secondary"}
                        size="small"
                        sx={{ fontSize: '0.7rem', height: 20 }}
                    />
                    <Chip
                        label={`${Math.round(zoomLevel * 100)}%`}
                        size="small"
                        sx={{ fontSize: '0.7rem', height: 20 }}
                    />
                    {isPending && <CircularProgress size={14} />}
                </Box>

                <Box sx={{ display: 'flex', gap: 1, alignItems: 'center' }}>
                    <Box sx={{ display: 'flex', gap: 0.5, alignItems: 'center' }}>
                        <IconButton size="small" onClick={handleZoomOut} disabled={zoomLevel <= 0.3}>
                            <Typography fontSize="0.9rem">-</Typography>
                        </IconButton>

                        <FormControl size="small" sx={{ minWidth: 80 }}>
                            <Select
                                value={zoomLevel}
                                onChange={(e) => setZoomLevel(Number(e.target.value))}
                                sx={{ fontSize: '0.75rem', height: 30 }}
                            >
                                {zoomOptions.map(option => (
                                    <MenuItem key={option.label} value={option.value} sx={{ fontSize: '0.75rem' }}>
                                        {option.label}
                                    </MenuItem>
                                ))}
                            </Select>
                        </FormControl>

                        <IconButton size="small" onClick={handleZoomIn} disabled={zoomLevel >= 2}>
                            <Typography fontSize="0.9rem">+</Typography>
                        </IconButton>

                        <IconButton size="small" onClick={handleZoomReset}>
                            <Rotate90DegreesCcw fontSize="small" />
                        </IconButton>
                    </Box>

                    <Divider orientation="vertical" flexItem />

                    <IconButton size="small"
                        onClick={() => setViewMode('preview')}
                        color={viewMode === 'preview' ? 'primary' : 'default'}>
                        <Pageview fontSize="small" />
                    </IconButton>

                    <IconButton size="small"
                        onClick={() => setViewMode('print')}
                        color={viewMode === 'print' ? 'primary' : 'default'}>
                        <Print fontSize="small" />
                    </IconButton>

                    <Divider orientation="vertical" flexItem />

                    <Button
                        size="small"
                        variant="contained"
                        startIcon={<Description fontSize="small" />}
                        onClick={() => handleExport('docx')}
                        disabled={exporting && exportFormat === 'docx'}
                        sx={{ fontSize: '0.7rem', minWidth: 70, backgroundColor: '#032F5D' }}
                    >
                        DOCX
                    </Button>

                    <Button
                        size="small"
                        variant="contained"
                        startIcon={<PictureAsPdf fontSize="small" />}
                        onClick={() => handleExport('pdf')}
                        disabled={exporting && exportFormat === 'pdf'}
                        color="error"
                        sx={{ fontSize: '0.7rem', minWidth: 60 }}
                    >
                        PDF
                    </Button>
                </Box>
            </Box>

            <Box sx={{
                flex: 1,
                overflowX: 'hidden',
                overflowY: 'auto',
                display: 'flex',
                justifyContent: 'center',
                padding: viewMode === 'preview' ? '20px 0' : '0',
                backgroundColor: '#f5f5f5',
                minHeight: viewMode === 'preview' ? 'auto' : '100%'
            }}>
                <Box
                    ref={reportPagesRef}
                    id="report-pages"
                    sx={{
                        display: 'flex',
                        flexDirection: 'column',
                        alignItems: 'center',
                        gap: viewMode === 'preview' ? (zoomLevel >= 1 ? 4 : 2) : 0,
                        transform: viewMode === 'preview' ? `scale(${zoomLevel})` : 'none',
                        transformOrigin: 'top center',
                        transition: 'transform 0.2s ease, gap 0.2s ease',
                        width: 'auto',
                        height: viewMode === 'preview' ? 'auto' : '100%',
                        marginTop: viewMode === 'preview' ? '20px' : '0',
                        marginBottom: viewMode === 'preview' ? '40px' : '0',
                        paddingBottom: viewMode === 'preview' ? '0' : '0',
                    }}
                >
                    <Page id="cover-page" pageNumber={pageNumbers['cover-page']?.start}>
                        <Box
                            sx={{
                                height: '100%',
                                display: 'flex',
                                flexDirection: 'column',
                                justifyContent: 'center',
                                alignItems: 'center',
                                px: 4
                            }}
                        >
                            <Box
                                sx={{
                                    display: 'flex',
                                    justifyContent: 'space-between',
                                    width: '100%',
                                    mb: 4,
                                    alignItems: 'center'
                                }}
                            >
                                <Box
                                    sx={{
                                        width: '180px',
                                        height: '80px',
                                        display: 'flex',
                                        alignItems: 'center',
                                        justifyContent: 'center',
                                        overflow: 'hidden'
                                    }}
                                >
                                    <img
                                        src={logo}
                                        alt="Atlas Logo"
                                        style={{
                                            maxWidth: '100%',
                                            maxHeight: '100%',
                                            objectFit: 'contain'
                                        }}
                                    />
                                </Box>

                                {bank?.logo_url && (
                                    <Box
                                        sx={{
                                            width: '180px',
                                            height: '80px',
                                            display: 'flex',
                                            alignItems: 'center',
                                            justifyContent: 'center',
                                            overflow: 'hidden'
                                        }}
                                    >
                                        <img
                                            src={`${config.backendURL}${bank.logo_url}`}
                                            alt={`${bank.name} Logo`}
                                            style={{
                                                maxWidth: '100%',
                                                maxHeight: '100%',
                                                objectFit: 'contain'
                                            }}
                                        />
                                    </Box>
                                )}
                            </Box>

                            <Typography
                                variant="h4"
                                sx={{ fontWeight: 'bold', mb: 1, textAlign: 'center' }}
                            >
                                Lenders Independent Engineers Report
                            </Typography>

                            <Typography
                                variant="h6"
                                sx={{ color: 'text.secondary', mb: 2, textAlign: 'center' }}
                            >
                                a project of
                            </Typography>

                            <Typography
                                variant="h3"
                                sx={{ fontWeight: 'bold', mb: 4, textAlign: 'center' }}
                            >
                                {project?.name}
                            </Typography>

                            {project?.cover_img && (
                                <Box
                                    sx={{
                                        width: '100%',
                                        maxHeight: '280px',
                                        my: 4,
                                        display: 'flex',
                                        justifyContent: 'center',
                                        alignItems: 'center',
                                        overflow: 'hidden'
                                    }}
                                >
                                    <img
                                        src={`${config.backendURL}${project.cover_img}`}
                                        alt="Project Cover"
                                        style={{
                                            maxWidth: '100%',
                                            maxHeight: '280px',
                                            objectFit: 'contain',
                                            borderRadius: '8px'
                                        }}
                                    />
                                </Box>
                            )}

                            <Box sx={{ textAlign: 'center', mt: 4 }}>
                                <Typography variant="h6" sx={{ color: 'text.secondary' }}>
                                    For
                                </Typography>

                                <Typography variant="h4" sx={{ fontWeight: 'bold', my: 2 }}>
                                    {bank?.name}
                                </Typography>

                                <Typography variant="h6" sx={{ color: 'text.secondary' }}>
                                    By
                                </Typography>

                                <Typography variant="h5" sx={{ fontWeight: 'bold', my: 2 }}>
                                    Adas Financial Research & Consulting Pvt. Ltd
                                </Typography>

                                <Typography variant="subtitle1" sx={{ mb: 1 }}>
                                    Aditya Trade Center, Ameerpet, Hyderabad
                                </Typography>

                                <Typography variant="h6" sx={{ color: 'primary.main', mb: 1 }}>
                                    <a
                                        href="https://www.atlasfin.in"
                                        target="_blank"
                                        rel="noopener noreferrer"
                                    >
                                        www.atlasfin.in
                                    </a>
                                </Typography>

                                <Typography variant="h6" sx={{ color: 'text.secondary' }}>
                                    {formatMonthYear(project?.created_at)}
                                </Typography>
                            </Box>
                        </Box>
                    </Page>

                    <Page id="disclaimer-page" pageNumber={pageNumbers['disclaimer-page']?.start}>
                        <Box sx={{ height: '100%', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
                            <Typography variant="h4" sx={{ fontWeight: 'bold', mb: 4, textAlign: 'center', color: '#032F5D' }}>
                                Disclaimer
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                This Lenders' Independent Engineer (LIE) Report as of {formatMonthYear(project?.created_at)} has been prepared at the request from Canara Bank, Large Corporate Branch, Punjagutta, Hyderabad. The report contains proprietary and confidential information.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                This LIE report has been prepared by M/s ATLAS Financial Research and Consulting Pvt Ltd., based
                                on information provided by the management, staff of the M/s Vasavi Group LLP and independent
                                study conducted by M/s ATLAS Financial Research and Consulting Pvt Ltd.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                This report provides progress of project as on date of site visit and is based on physical and financial
                                information submitted by company.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                M/s ATLAS Financial Research and Consulting Pvt follows ethical practices in the discharge of its
                                professional services and amongst others, as part of such ethical practices, it follows the general rules
                                relating to honesty, competence, and confidentiality, and attempts to provide the most current,
                                complete, and accurate information as possible within the limitations of available finance, time
                                constraint and other practical difficulties relating thereto and arising as a consequence thereof.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                It is further informed; no responsibility is accepted by Atlas financials and/or its affiliates and/or its
                                Directors, employees/officers for this report or for any direct or consequential loss arising from any
                                use of the information, statements, or forecasts in the Report.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                This report is furnished on a strictly confidential basis and is for the sole use of the company and all
                                the lenders of the company only. Neither this report nor the information contained herein may be
                                reproduced or passed to any person or used for any purpose other than stated above. By accepting a
                                copy of this LIE report, the recipients accept the terms of this notice, which forms an integral part of
                                this LIE report.
                            </Typography>
                        </Box>
                    </Page>

                    <Page id="about-page" pageNumber={pageNumbers['about-page']?.start}>
                        <Box sx={{ height: '100%', display: 'flex', flexDirection: 'column', justifyContent: 'center' }}>
                            <Typography variant="h4" sx={{ fontWeight: 'bold', mb: 4, textAlign: 'center', color: '#032F5D' }}>
                                About Atlas Financial Research & Consulting (P) Ltd
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                Atlas Financial Research & Consulting (P) Ltd. is a distinguished company specializing in financial solutions and technical services essential for our client business's triumphant journey. With over a decade of experience, the Atlas team consists of esteemed professionals, including former senior executives from both public and private sector banks and seasoned technocrats. This wealth of expertise positions Atlas perfectly to cater to the diverse needs of the client.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                With corporate headquarters based in Hyderabad, Atlas has established a robust presence in prominent
                                cities across India, such as Mumbai, Bangalore, Delhi, Kolkata, Bhopal, and Andhra Pradesh. Atlas has
                                completed over 510 assignments throughout its journey, showcasing its proficiency in TEV Studies and
                                LIE assignments across various industries. This track record exemplifies Atlas dedication to deliver
                                outstanding results for its clients.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                In addition to TEV and LIE services, Atlas also provides an array of services encompassing Detailed
                                Project Reports, Pre-Feasibility Reports for land allocation, Project Consulting and Advisory, as well as
                                Business Consulting Assignments including Asset Valuation, Enterprise Valuation, Sustainability
                                Reporting and M&A advisory.
                            </Typography>
                            <Typography variant="body1" sx={{ lineHeight: 1.8, mb: 2.5 }}>
                                Atlas takes immense pride in its esteemed associations with major banks in the country, such as State
                                Bank of India, Canara Bank, Bank of Baroda, Union Bank of India, Central Bank of India, Indian Bank,
                                Bank of Maharashtra, National Bank for Financing Infrastructure and Development, YES Bank, UCO
                                Bank, and others. This recognition speaks volumes about its standing in the financial industry and the
                                trust placed in us by these reputable institutions
                            </Typography>
                        </Box>
                    </Page>

                    {tocPages.length > 0 ? (
                        tocPages.map((tocPage) => (
                            <Page
                                key={`toc-page-${tocPage.pageNumber}`}
                                id={`toc-page-${tocPage.pageNumber}`}
                                pageNumber={pageNumbers[`toc-page-${tocPage.pageNumber}`]?.start}
                                className="toc-page"
                            >
                                <TOCPage
                                    items={tocPage.items}
                                    pageNumbers={pageNumbers}
                                    scrollToPage={scrollToPage}
                                    pageNumber={tocPage.pageNumber}
                                    totalPages={tocTotalPages}
                                />
                            </Page>
                        ))
                    ) : (
                        <Page id="toc-page" pageNumber={pageNumbers['toc-page']?.start}>
                            <Box sx={{ height: '100%' }}>
                                <TableOfContents
                                    checkedByGroup={checkedByGroup}
                                    pageNumbers={pageNumbers}
                                    scrollToPage={scrollToPage}
                                />
                            </Box>
                        </Page>
                    )}

                    {checkedByGroup.map((group, groupIndex) => {
                        if (!group?.checked) return null;

                        const sectionNumber = sectionNumbering[group.id] || (groupIndex + 1).toString();
                        const sectionPages = splitContentIntoPages(group.content || [], `section-${group.id}`);

                        const hasSectionComments = permissionService.hasPermission("comments");

                        return (
                            <React.Fragment key={group.id}>
                                {sectionPages.pages.map((pageContent, pageIndex) => {
                                    const isLastSectionPage = pageIndex === sectionPages.pages.length - 1;

                                    return (
                                        <ContentPage
                                            key={`section-${group.id}-page-${pageIndex + 1}`}
                                            pageId={`section-${group.id}-page-${pageIndex + 1}`}
                                            pageNumber={pageNumbers[`section-${group.id}-page-1`]?.start + pageIndex}
                                            title={pageIndex === 0 ? `${sectionNumber}. ${group.title}` : `${sectionNumber}. ${group.title} (continued)`}
                                            sectionNumber=""
                                        >
                                            <Box sx={{ lineHeight: 1.6 }}>
                                                {hasSectionComments && pageIndex === 0 && (
                                                    <Box sx={{ display: 'flex', justifyContent: 'flex-end', mb: 2 }}>
                                                        <IconButton
                                                            size="small"
                                                            onClick={() => {
                                                                if (!group.main_id) return;
                                                                handleOpenCommentModal("section", group.main_id);
                                                            }}
                                                            sx={{ color: '#666' }}
                                                            title="Add comment to this section"
                                                        >
                                                            <CommentOutlined fontSize="small" />
                                                        </IconButton>
                                                    </Box>
                                                )}

                                                {pageContent.map((content, contentIndex) => (
                                                    <Box key={contentIndex} sx={{ mb: 3 }}>
                                                        {renderOlResult(
                                                            `section-${group.id}-content-${contentIndex}`,
                                                            content,
                                                            group.main_id || '',
                                                            contentIndex,
                                                            pageContent,
                                                            handleOpenCommentModal
                                                        )}
                                                    </Box>
                                                ))}

                                                {isLastSectionPage && hasSectionComments && group.main_id && (
                                                    <Box sx={{ mt: 4, pt: 3, borderTop: '1px solid #e5e7eb' }}>
                                                        <SectionComments
                                                            sectionId={group.main_id}
                                                            safeFetch={safeFetch}
                                                        />
                                                    </Box>
                                                )}
                                            </Box>
                                        </ContentPage>
                                    );
                                })}

                                {group.items.map((item, itemIndex) => {
                                    if (!item.checked) return null;

                                    const subsectionNumber = subsectionNumbering[item.id] || `${sectionNumber}.${itemIndex + 1}`;
                                    const itemPages = splitContentIntoPages(item.content || [], `subsection-${item.id}`);

                                    return (
                                        <React.Fragment key={item.id}>
                                            {itemPages.pages.map((pageContent, pageIndex) => {
                                                const hasWideTable = pageContent.some(content =>
                                                    content.content_type === 'table_horizontal' &&
                                                    content.headers &&
                                                    content.headers.length > 5
                                                );

                                                const isLastSubsectionPage = pageIndex === itemPages.pages.length - 1;
                                                const hasSubsectionComments = permissionService.hasPermission("comments");

                                                return (
                                                    <ContentPage
                                                        key={`subsection-${item.id}-page-${pageIndex + 1}`}
                                                        pageId={`subsection-${item.id}-page-${pageIndex + 1}`}
                                                        pageNumber={pageNumbers[`subsection-${item.id}-page-1`]?.start + pageIndex}
                                                        title={pageIndex === 0 ? `${subsectionNumber} ${item.label}` : `${subsectionNumber} ${item.label} (continued)`}
                                                        orientation={hasWideTable ? 'landscape' : 'portrait'}
                                                        isSubsection
                                                        sectionNumber=""
                                                    >
                                                        <Box sx={{
                                                            lineHeight: 1.6,
                                                            width: '100%',
                                                            height: '100%',
                                                            display: 'flex',
                                                            flexDirection: 'column'
                                                        }}>
                                                            {hasSubsectionComments && pageIndex === 0 && (
                                                                <Box sx={{ display: 'flex', justifyContent: 'flex-end', mb: 2 }}>
                                                                    <IconButton
                                                                        size="small"
                                                                        onClick={() => {
                                                                            if (!item.main_id) return;
                                                                            handleOpenCommentModal("subsection", item.main_id);
                                                                        }}
                                                                        sx={{ color: '#666' }}
                                                                        title="Add comment to this subsection"
                                                                    >
                                                                        <CommentOutlined fontSize="small" />
                                                                    </IconButton>
                                                                </Box>
                                                            )}

                                                            {pageContent.map((content, contentIndex) => (
                                                                <Box key={contentIndex} sx={{
                                                                    mb: 3,
                                                                    width: '100%',
                                                                    overflowX: hasWideTable ? 'auto' : 'hidden'
                                                                }}>
                                                                    {renderOlResult(
                                                                        `subsection-${item.id}-content-${contentIndex}`,
                                                                        content,
                                                                        item.main_id || '',
                                                                        contentIndex,
                                                                        pageContent,
                                                                        handleOpenCommentModal
                                                                    )}
                                                                </Box>
                                                            ))}

                                                            {isLastSubsectionPage && hasSubsectionComments && item.main_id && (
                                                                <Box sx={{ mt: 'auto', pt: 3, borderTop: '1px solid #e5e7eb' }}>
                                                                    <SubsectionComments
                                                                        subsectionId={item.main_id}
                                                                        safeFetch={safeFetch}
                                                                    />
                                                                </Box>
                                                            )}
                                                        </Box>
                                                    </ContentPage>
                                                );
                                            })}
                                        </React.Fragment>
                                    );
                                })}
                            </React.Fragment>
                        );
                    })}

                    <Page id="annexures-page" pageNumber={pageNumbers['annexures-page']?.start}>
                        <Box sx={{ height: '100%', textAlign: 'center' }}>
                            <Typography variant="h4" sx={{ fontWeight: 'bold', mb: 3, color: '#2563EB' }}>
                                Annexures 2
                            </Typography>
                            <Typography variant="h6" sx={{ color: 'text.secondary', mb: 4 }}>
                                just a subtitle
                            </Typography>
                            <Box sx={{
                                display: 'grid',
                                gridTemplateColumns: 'repeat(4, 1fr)',
                                gap: 2,
                                mt: 4,
                                width: '100%'
                            }}>
                                {annexurePhotos.length === 0 ? (
                                    <Typography variant="body2" color="text.secondary">
                                        No site visit photos available
                                    </Typography>
                                ) : (
                                    annexurePhotos.map((photo, index) => (
                                        <Box
                                            key={photo.id || index}
                                            sx={{
                                                width: '100%',
                                                aspectRatio: '1 / 1',
                                                overflow: 'hidden',
                                                borderRadius: '8px',
                                                boxShadow: 2
                                            }}
                                        >
                                            <img
                                                src={
                                                    photo.file_url.startsWith("http")
                                                        ? photo.file_url
                                                        : `${config.backendURL}${photo.file_url}`
                                                }
                                                alt={`Project photo ${index + 1}`}
                                                loading="lazy"
                                                decoding="async"
                                                style={{
                                                    width: '100%',
                                                    height: '100%',
                                                    objectFit: 'cover'
                                                }}
                                            />
                                        </Box>
                                    ))
                                )}
                            </Box>
                        </Box>
                    </Page>
                </Box>
            </Box>

            <Dialog open={exporting} maxWidth="sm" fullWidth>
                <DialogTitle>
                    Exporting Report...
                </DialogTitle>
                <DialogContent>
                    <Box sx={{ display: 'flex', flexDirection: 'column', alignItems: 'center', p: 3 }}>
                        <CircularProgress size={60} />
                        <Typography variant="body1" sx={{ mt: 2 }}>
                            {exportFormat === 'docx'
                                ? 'Generating Microsoft Word document...'
                                : 'Generating PDF document...'}
                        </Typography>
                        <Typography variant="caption" color="text.secondary" sx={{ mt: 1 }}>
                            This may take a moment depending on the report size.
                        </Typography>
                    </Box>
                </DialogContent>
            </Dialog>
        </Box>
    );
});

/* ------------------------------------------------- Main Component ------------------------------------------------- */
export default function LIEReport() {
    const { id } = useParams<{ id: string }>();
    const { userId } = useAuth();

    // State
    const [userInfo, setUserInfo] = useState<UserDataInfo | null>(null);
    const [role, setRole] = useState<RoleDataInfo | null>(null);
    const [stage, setStage] = useState<string>("Draft");
    const [currentCheck, setCurrentCheck] = useState<string>('');
    const [isAsideMenuOpen, setIsAsideMenuOpen] = useState<boolean>(false);
    const [newSectionName, setNewSectionName] = useState<string>("");

    // Data groups state
    const [persist, setPersist] = useState<{ groups: Group[]; olResults: Record<string, OlgptAny[]> }>({
        groups: [],
        olResults: {}
    });
    const [groups, setGroups] = useState<Group[]>(persist.groups);
    const [olResults, setOlResults] = useState<Record<string, OlgptAny[]>>(persist.olResults);
    const [statusLogs, setStatusLogs] = useState<ReportStatusLog[]>([]);
    const [summaryPopupOpen, setSummaryPopupOpen] = useState(false);
    const [summaryPayload, setSummaryPayload] = useState<{
        project_id: string;
        report_version_id: string;
        sections: {
            id: string;
            title: string;
            isNew: boolean;
            selected: boolean;
        }[];
        subsections: {
            id: string;
            label: string;
            isNew: boolean;
            section_id: string;
            selected: boolean;
        }[];
    } | null>(null);
    const [openStatus, setOpenStatus] = useState(false);
    const [loadingStatus, setLoadingStatus] = useState(false);
    const [processingJobs, setProcessingJobs] = useState<ProcessingJobLog[]>([]);
    const [confirmStage, setConfirmStage] = useState<string | null>(null);
    const [showBulkQueryDialog, setShowBulkQueryDialog] = useState(false);
    const [pendingBulkGeneration, setPendingBulkGeneration] = useState<{
        sections: any[];
        subsections: any[];
    } | null>(null);
    // const [reviewingSection, setReviewingSection] = useState<{ jobId: string, subsectionId: string, sectionName: string } | null>(null);
    const [openPreview, setOpenPreview] = useState(false);
    const [previewContent, setPreviewContent] = useState<string>('');
    const [generatingMap, setGeneratingMap] = useState<Record<string, boolean>>({});
    const [lastGenerationPayload, setLastGenerationPayload] = useState<any>(null);
    const [reviewingSection, setReviewingSection] = useState<any>(null);
    const [regenQuery, setRegenQuery] = useState("");
    const [snackbar, setSnackbar] = useState<{
        open: boolean;
        message: string;
        severity: 'error' | 'warning' | 'success';
    }>({
        open: false,
        message: '',
        severity: 'warning'
    });

    // Custom hooks for data fetching
    const {
        project,
        bank,
        industry,
        client,
        loading: projectLoading,
        error: projectError,
        reload: reloadProject
    } = useProjectData(id || '');

    const {
        allSection,
        allSubSection,
        reportV,
        projectRS,
        checklists,
        reload: reloadReportData,
        setAllSection,
        setAllSubSection,
        setReportV
    } = useReportData(project);

    const updateReportStatus = useCallback(
        async (nextStatus: string) => {
            if (!reportV?.id) return;

            try {
                await fetch(
                    `${config.backendURL}/api/report-versions/${reportV.id}/status`,
                    {
                        method: "PUT",
                        credentials: "include",
                        headers: { "Content-Type": "application/json" },
                        body: JSON.stringify({ status: nextStatus }),
                    }
                );

                await reloadReportData();
                await loadStatusLogs();
            } catch (err) {
                console.error("Failed to update report status:", err);
            }
        },
        [reportV?.id, reloadReportData]
    );

    const fetchProcessingJobs = async () => {
        setLoadingStatus(true);
        try {
            const res = await fetch(
                `${config.backendURL}/api/project-report/status-log?project_id=${project?.id}`,
                { credentials: 'include' }
            );

            const text = await res.text();
            // console.log("RAW response:", text);

            let data;
            try {
                data = JSON.parse(text);
            } catch {
                console.error("Not JSON. Backend returned HTML or text.");
                setProcessingJobs([]);
                return;
            }

            // Process the data to handle report_json structure
            const processedJobs: ProcessingJobLog[] = Array.isArray(data)
                ? data.map((job: ProcessingJobLog) => {
                    // If report_json exists and has toc array
                    if (job.report_json && job.report_json.toc && Array.isArray(job.report_json.toc)) {
                        return {
                            ...job,
                            tocItems: job.report_json.toc
                        };
                    }
                    // If report_json is directly an array (fallback)
                    else if (job.report_json && Array.isArray(job.report_json)) {
                        return {
                            ...job,
                            tocItems: job.report_json
                        };
                    }
                    // If no valid report_json structure
                    else {
                        return {
                            ...job,
                            tocItems: []
                        };
                    }
                }) : [];

            setProcessingJobs(processedJobs);
        } catch (err) {
            console.error("Failed to fetch processing jobs", err);
            setProcessingJobs([]);
        } finally {
            setLoadingStatus(false);
        }
    };

    // Function to handle approve/reject
    const handleReviewSection = async (jobId: string, projectReportSectionsId: string, action: 'approved' | 'rejected') => {
        try {
            const res = await fetch(`${config.backendURL}/api/reports/review`, {
                method: "POST",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    job_id: jobId,
                    project_report_sections_id: projectReportSectionsId,
                    action: action
                })
            });

            if (res.ok) {
                // Refresh the jobs list
                await fetchProcessingJobs();
                // Refresh report data if approved
                if (action === 'approved') {
                    await reloadReportData();
                }
                console.log(`Section ${action} successfully!`);
                setReviewingSection(null);
            } else {
                throw new Error('Failed to update section');
            }
        } catch (error) {
            console.error("Error reviewing section:", error);
            console.log("Failed to update section. Please try again.");
        }
    };

    const safeFetch = useSafeFetch();
    const sensors = useSensors(useSensor(PointerSensor));
    const prevExpandedRef = useRef<Record<string, boolean>>({});
    const { print } = usePrint({
        targetId: 'report-pages',
        pageMarginMM: 15,
        keepColors: true,
    });

    // Load user info and role
    useEffect(() => {
        if (!userId) return;

        const loadUserInfo = async () => {
            try {
                const info = await safeFetch(`/api/users/${userId}`);
                setUserInfo(info);

                if (info.role_id) {
                    const roleData = await safeFetch(`/api/roles/${info.role_id}`);
                    setRole(roleData);
                }
            } catch (e) {
                console.error("Failed to fetch user info:", e);
            }
        };

        loadUserInfo();
    }, [userId, safeFetch]);

    const normalizeProjectTypes = (value: any): string[] => {
        if (!value) return [];

        if (Array.isArray(value)) {
            return value.map(v => String(v).trim());
        }

        if (typeof value === "number") {
            return [String(value)];
        }

        if (typeof value === "string") {
            return value
                .split(',')
                .map(v => v.trim())
                .filter(Boolean);
        }

        try {
            const parsed = JSON.parse(value);
            return normalizeProjectTypes(parsed);
        } catch {
            return [];
        }
    };

    const supportsProjectType = (
        templateProjectType: any,
        projectType: any
    ): boolean => {
        const templateTypes = normalizeProjectTypes(templateProjectType);
        const projectTypes = normalizeProjectTypes(projectType);

        return projectTypes.some(pt => templateTypes.includes(pt));
    };

    // Update groups when data changes
    useEffect(() => {
        if (projectRS && reportV && allSection && allSubSection) {
            const filteredPRS = projectRS.filter(prs => prs.report_version_id === reportV.id);

            const eligibleSections = allSection.filter(sec =>
                supportsProjectType(sec.project_type, project?.project_type)
            );

            const sortedSections = [...eligibleSections].sort((a, b) => {
                const aPrs = filteredPRS.find(
                    p =>
                        p.section_template_id === a.id &&
                        (!p.subsection_template_id || p.subsection_template_id === "")
                );

                const bPrs = filteredPRS.find(
                    p =>
                        p.section_template_id === b.id &&
                        (!p.subsection_template_id || p.subsection_template_id === "")
                );

                return (
                    Number(aPrs?.order_index ?? a.order_suggestion ?? 0) -
                    Number(bPrs?.order_index ?? b.order_suggestion ?? 0)
                );
            });

            const newGroups: Group[] = sortedSections.map((sec, sIdx) => {
                const groupId = sec?.id ?? `section-${sIdx}`;
                const groupTitle = sec?.name ?? `Section ${sIdx + 1}`;

                const items: Task[] = allSubSection
                    .filter(sub =>
                        sub.section_template_id === sec.id &&
                        supportsProjectType(sub.project_type, project?.project_type)
                    )
                    .map((sub, i) => {
                        const prs = filteredPRS.find(p =>
                            p.section_template_id === sec.id &&
                            p.subsection_template_id === sub.id
                        );

                        return {
                            id: sub.id!,
                            label: sub.name!,
                            checked: prs?.checked,
                            order_sub_index: Number(prs?.order_sub_index ?? i),
                            content: prs?.content ? normalizeContent(prs.content) : [],
                            main_id: prs?.id ?? null,
                        };
                    })
                    .sort((a, b) => a.order_sub_index - b.order_sub_index);


                const groupPrs = filteredPRS.find(
                    p =>
                        p.section_template_id === sec.id &&
                        (p.subsection_template_id === null || p.subsection_template_id === "")
                );
                const normalized = normalizeContent(groupPrs?.content);

                return {
                    id: groupId,
                    title: groupTitle,
                    expanded: groups.find(g => g.id === groupId)?.expanded ?? false,
                    checked:
                        groupPrs?.checked ??
                        items.some(item => item.checked),
                    order_index: groupPrs?.order_index ?? sIdx.toString(),
                    items,
                    content: Array.isArray(normalized)
                        ? normalized
                        : normalized
                            ? [normalized]
                            : [],
                    main_id: groupPrs?.id ?? null,
                };
            });

            setPersist(prev => ({ ...prev, groups: newGroups }));
        }
    }, [projectRS, reportV, allSection, allSubSection, groups]);

    // Sync local state with persist
    useEffect(() => {
        setGroups(persist.groups);
    }, [persist.groups]);

    useEffect(() => {
        setPersist({ groups, olResults });
    }, [groups, olResults]);

    // Update stage when reportV changes
    useEffect(() => {
        if (reportV?.status) {
            setStage(reportV.status);
        }
    }, [reportV]);

    // Drag and drop handlers
    const findGroupByItemId = useCallback(
        (itemId: UniqueIdentifier) => groups.find(g => g.items.some(t => t.id === itemId)),
        [groups]
    );

    const getGroupIndex = useCallback(
        (groupId: UniqueIdentifier) => groups.findIndex(g => g.id === groupId),
        [groups]
    );

    const handleDragStart = useCallback((e: DragStartEvent) => {
        if (e.active.data.current?.type === 'container') {
            const snap: Record<string, boolean> = {};
            for (const g of groups) snap[g.id] = !!g.expanded;
            prevExpandedRef.current = snap;
            setGroups(prev => prev.map(g => ({ ...g, expanded: false })));
        }
    }, [groups]);

    const restoreExpandedIfNeeded = useCallback(() => {
        const snap = prevExpandedRef.current;
        if (snap && Object.keys(snap).length) {
            setGroups(prev => prev.map(g => ({ ...g, expanded: snap[g.id] ?? g.expanded })));
            prevExpandedRef.current = {};
        }
    }, []);

    const handleDragOver = useCallback((event: DragOverEvent) => {
        const { active, over } = event;
        if (!over) return;

        if (active.data.current?.type === 'item') {
            const fromGroup = findGroupByItemId(active.id);
            const overIsItem = over.data.current?.type === 'item';
            const toGroup = overIsItem ? findGroupByItemId(over.id) : groups.find(g => g.id === over.id);

            if (!fromGroup || !toGroup || fromGroup.id === toGroup.id) return;

            const fromIdx = groups.findIndex(g => g.id === fromGroup.id);
            const toIdx = groups.findIndex(g => g.id === toGroup.id);
            const fromItems = [...groups[fromIdx].items];
            const movingIndex = fromItems.findIndex(t => t.id === active.id);

            if (movingIndex < 0) return;

            const [moving] = fromItems.splice(movingIndex, 1);
            const toItems = [...groups[toIdx].items];
            const insertIndex = overIsItem ? toItems.findIndex(t => t.id === over.id) : toItems.length;

            toItems.splice(insertIndex < 0 ? toItems.length : insertIndex, 0, moving);

            const next = [...groups];
            next[fromIdx] = { ...next[fromIdx], items: fromItems };
            next[toIdx] = { ...next[toIdx], items: toItems };
            setGroups(next);
        }
    }, [groups, findGroupByItemId]);

    const saveToBackend = useCallback(async (
        changedSections: any[] = [],
        changedSubSections: any[] = [],
        isParent = false,
        isChild = false
    ) => {
        if (!project?.version_id) return;

        try {
            const [section_temps, sub_section_temps] = simplifyReportData(groups);
            const payload = {
                report_version_id: project.version_id,
                sections: isParent ? (changedSections.length ? changedSections : section_temps) : [],
                subsections: isChild ? (changedSubSections.length ? changedSubSections : sub_section_temps) : [],
            };

            const uploadRes = await fetch(`${config.backendURL}/api/project-report-sections/update-order`, {
                method: "POST",
                body: JSON.stringify(payload),
                credentials: "include",
                headers: { "Content-Type": "application/json" },
            });

            if (!uploadRes.ok) {
                throw new Error(`Failed to update order`);
            }
        } catch (error) {
            console.error("Error saving to backend:", error);
        }
    }, [groups, project]);


    const handleDragEnd = useCallback((event: DragEndEvent) => {
        const { active, over } = event;

        if (!over) {
            restoreExpandedIfNeeded();
            return;
        }

        const isParent = active.data.current?.type === 'container';
        const isChild = active.data.current?.type === 'item';

        if (isParent && over.data.current?.type === 'container') {
            const from = getGroupIndex(active.id);
            const to = getGroupIndex(over.id);

            if (from !== -1 && to !== -1 && from !== to) {
                setGroups(prev => {
                    const newGroups = arrayMove(prev, from, to).map((g, i) => ({
                        ...g,
                        order_index: i.toString(),
                    }));

                    const [section_temps] = simplifyReportData(newGroups);
                    saveToBackend(section_temps, [], true, false);
                    return newGroups;
                });
            }

            restoreExpandedIfNeeded();
            return;
        }

        if (isChild && over.data.current?.type === 'item') {
            const fromGroup = findGroupByItemId(active.id);
            const toGroup = findGroupByItemId(over.id);

            if (!fromGroup || !toGroup) {
                restoreExpandedIfNeeded();
                return;
            }

            if (fromGroup.id === toGroup.id) {
                const gi = getGroupIndex(fromGroup.id);
                const items = groups[gi].items;
                const oldIndex = items.findIndex(t => t.id === active.id);
                const newIndex = items.findIndex(t => t.id === over.id);

                if (oldIndex !== -1 && newIndex !== -1 && oldIndex !== newIndex) {
                    setGroups(prev => {
                        const next = [...prev];
                        const reorderedItems = arrayMove(items, oldIndex, newIndex).map((t, i) => ({
                            ...t,
                            order_sub_index: i,
                        }));
                        next[gi] = { ...next[gi], items: reorderedItems };

                        const [, sub_section_temps] = simplifyReportData([next[gi]]);
                        saveToBackend([], sub_section_temps, false, true);
                        return next;
                    });
                }
            } else {
                setGroups(prev => {
                    const newGroups = [...prev];
                    const fromGroupIdx = newGroups.findIndex(g => g.id === fromGroup.id);
                    const toGroupIdx = newGroups.findIndex(g => g.id === toGroup.id);

                    const fromItems = [...newGroups[fromGroupIdx].items];
                    const toItems = [...newGroups[toGroupIdx].items];

                    const movingIdx = fromItems.findIndex(t => t.id === active.id);
                    if (movingIdx === -1) return prev;

                    const [movingItem] = fromItems.splice(movingIdx, 1);
                    const targetIdx = toItems.findIndex(t => t.id === over.id);
                    const insertIdx = targetIdx === -1 ? toItems.length : targetIdx;
                    toItems.splice(insertIdx, 0, movingItem);

                    newGroups[fromGroupIdx] = {
                        ...newGroups[fromGroupIdx],
                        items: fromItems.map((item, i) => ({ ...item, order_sub_index: i })),
                    };
                    newGroups[toGroupIdx] = {
                        ...newGroups[toGroupIdx],
                        items: toItems.map((item, i) => ({ ...item, order_sub_index: i })),
                    };

                    const [, sub_section_temps] = simplifyReportData([newGroups[toGroupIdx]]);
                    saveToBackend([], sub_section_temps, false, true);
                    return newGroups;
                });
            }

            restoreExpandedIfNeeded();
        }
    }, [groups, findGroupByItemId, getGroupIndex, restoreExpandedIfNeeded, saveToBackend]);

    // Group and task handlers
    const toggleGroup = useCallback((id: string) => {
        setGroups(prev => prev.map(g => g.id === id ? { ...g, expanded: !g.expanded } : g));
    }, []);

    const updateCheckedStatus = useCallback(async (
        type: 'section' | 'subsection',
        id: string,
        currentChecked: boolean
    ) => {
        if (!project?.version_id) return;

        const payload =
            type === 'section'
                ? (() => {
                    const group = groups.find(g => g.id === id);
                    return {
                        report_version_id: project.version_id,
                        section_template_id: id,
                        checked: !currentChecked,
                        order_index: Number(group?.order_index),
                    };
                })()
                : (() => {
                    const group = groups.find(g => g.items.some(t => t.id === id));
                    const task = group?.items.find(t => t.id === id);
                    return {
                        report_version_id: project.version_id,
                        subsection_template_id: id,
                        checked: !currentChecked,
                        order_sub_index: Number(task?.order_sub_index),
                        section_template_id: group?.id,
                    };
                })();

        console.log("updating main section", payload);

        await fetch(
            `${config.backendURL}/api/project-report-sections/update-checked`,
            {
                method: "POST",
                body: JSON.stringify(payload),
                credentials: "include",
                headers: { "Content-Type": "application/json" },
            }
        );

        await reloadReportData();
    }, [groups, project, reloadReportData]);

    const toggleGroupAllChecked = useCallback((id: string) => {
        const group = groups.find(g => g.id === id);
        if (!group) return;

        updateCheckedStatus('section', id, !!group.checked);
        setGroups(prev => prev.map(g => g.id === id ? { ...g, checked: !g.checked } : g));
    }, [groups, updateCheckedStatus]);

    const toggleTaskChecked = useCallback((taskId: string) => {
        const group = groups.find(g => g.items.some(item => item.id === taskId));
        const task = group?.items.find(item => item.id === taskId);

        if (!task) return;

        updateCheckedStatus('subsection', taskId, !!task.checked);
        setGroups(prev => prev.map(g => {
            const idx = g.items.findIndex(t => t.id === taskId);
            if (idx === -1) return g;
            const items = [...g.items];
            items[idx] = { ...items[idx], checked: !items[idx]?.checked };
            return { ...g, items };
        }));
    }, [groups, updateCheckedStatus]);

    const handleAddSection = useCallback(async () => {
        if (!newSectionName.trim() || !project?.industry_id) return;

        const name = newSectionName.trim();
        setNewSectionName("");

        try {
            const newSection = await fetch(`${config.backendURL}/api/section-templates`, {
                method: "POST",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    industry_id: project.industry_id,
                    name,
                    default_content: JSON.stringify([{
                        content_type: "paragraph",
                        content: "Section content goes here..."
                    }]),
                    is_mandatory: false,
                    project_type: project.project_type,
                    order_suggestion: allSection.length + 1,
                    section_type: "custom",
                }),
            }).then(r => r.json());

            await fetch(`${config.backendURL}/api/project-report-sections/from-template`, {
                method: "POST",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                    report_version_id: reportV ? reportV.id : null,
                    section_template_id: newSection.id,
                }),
            });

            const updatedSections = await safeFetch(`/api/section-templates`);
            setAllSection(updatedSections);
        } catch (error) {
            console.error("Error adding section:", error);
        }
    }, [newSectionName, project, allSection.length, reportV, safeFetch, setAllSection]);

    // Individual generate content function
    const handleGenerateContent = useCallback(async (sectionId: string, subsectionId?: string, query?: string) => {
        if (!project || !reportV) return;

        const section = groups.find(g => g.id === sectionId);
        if (!section) return;

        const prsId =
            subsectionId
                ? projectRS.find(p =>
                    p.section_template_id === sectionId &&
                    p.subsection_template_id === subsectionId &&
                    p.report_version_id === reportV.id
                )?.id
                : section.main_id;

        if (prsId && generatingMap[prsId]) {
            console.warn("Generation already in progress. Ignored.");
            return;
        }

        const toc = [];

        if (subsectionId) {
            const subsection = section.items.find(i => i.id === subsectionId);
            if (!subsection) return;

            toc.push({
                section_id: sectionId,
                section_number: String(section.order_index ?? ""),
                section_name: section.title,
                subsection_id: subsectionId,
                query: query || "",
                subsection_number: String(subsection.order_sub_index ?? ""),
                subsection_name: subsection.label,
                project_report_sections_id: projectRS.find(p =>
                    p.section_template_id === sectionId &&
                    p.subsection_template_id === subsectionId &&
                    p.report_version_id === reportV.id
                )?.id ?? null,
            });
        } else {
            toc.push({
                section_id: sectionId,
                section_number: String(section.order_index ?? ""),
                section_name: section.title,
                subsection_id: "",
                query: query || "",
                subsection_number: "",
                subsection_name: "",
                project_report_sections_id: section.main_id ?? null,
            });
        }

        const payload = {
            project_id: project.id!,
            generation_type: "section",
            toc
        };

        console.log("📤 Enterprise OL-GPT REQUEST BODY:", payload);

        setLastGenerationPayload(payload);

        try {
            const res = await fetch(`${config.backendURL}/api/reports/generate`, {
                method: "POST",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            });
            console.log("Generate response:", res);

            if (!res.ok) {
                console.log("Entered !res.ok");
                let errorData: any = null;
                let errorText: any = null;

                try {
                    errorText = await res.text();
                    errorData = errorText ? JSON.parse(errorText) : null;
                    console.log("errorData:", errorData);
                } catch {
                    errorData = null;
                }

                // 🔔 HANDLE 409 HERE
                if (res.status === 409) {
                    const message =
                        errorData?.message ||
                        errorData?.reason ||
                        res.statusText ||
                        "Generation not allowed";

                    console.warn("Generation conflict:", message);

                    setSnackbar({
                        open: true,
                        message,
                        severity: "warning"
                    });

                    return null; // ⛔ STOP execution cleanly
                }

                throw new Error(`HTTP error! status: ${res.status}`);
            }

            const data = await res.json();
            console.log("RESPONSE BODY:", data);

            console.log(`Content generated successfully for ${subsectionId ? 'subsection' : 'section'}!`);

            await reloadReportData();

        } catch (error) {
            console.error("Failed to generate content:", error);
            console.log("Failed to generate content. Please try again.");
        }
    }, [project, reportV, groups, projectRS, reloadReportData]);

    const renderOlResult = useCallback(
        (
            key: string,
            content: OlgptAny,
            sectionId: string,
            index: number,
            fullContent: OlgptAny[],
            onOpenCommentModal?: (entityType: string, entityId: string) => void
        ) => {
            if (content.content_type === "paragraph") {
                const ParagraphEditor = () => {
                    const [editing, setEditing] = useState(false);
                    const [value, setValue] = useState(content.content);
                    const [saving, setSaving] = useState(false);

                    const save = async () => {
                        setSaving(true);
                        const updated = [...fullContent];
                        updated[index] = { ...content, content: value };

                        await fetch(
                            `${config.backendURL}/api/project-report-sections/${sectionId}`,
                            {
                                method: "PUT",
                                credentials: "include",
                                headers: { "Content-Type": "application/json" },
                                body: JSON.stringify({ content: updated }),
                            }
                        );

                        setSaving(false);
                        setEditing(false);
                    };

                    return (
                        <Box className="bg-white p-3 rounded-md" sx={{ position: 'relative' }}>
                            {!editing ? (
                                <>
                                    <Typography variant="body2" sx={{ lineHeight: 1.6, whiteSpace: 'pre-wrap' }}>
                                        {value}
                                    </Typography>
                                    <Button size="small" onClick={() => setEditing(true)} sx={{ mt: 1 }}>
                                        Edit
                                    </Button>
                                </>
                            ) : (
                                <>
                                    <TextField
                                        fullWidth
                                        multiline
                                        minRows={3}
                                        value={value}
                                        onChange={(e) => setValue(e.target.value)}
                                        sx={{ mb: 1 }}
                                    />
                                    <Stack direction="row" spacing={1} mt={1}>
                                        <Button size="small" onClick={save} disabled={saving}>
                                            Save
                                        </Button>
                                        <Button size="small" onClick={() => setEditing(false)}>
                                            Cancel
                                        </Button>
                                    </Stack>
                                </>
                            )}
                        </Box>
                    );
                };

                return <ParagraphEditor />;
            }

            if (content.content_type === "table_horizontal") {
                const TableEditor = () => {
                    const [editing, setEditing] = useState(false);
                    const [headers, setHeaders] = useState([...content.headers]);
                    const [rows, setRows] = useState([...content.rows]);

                    const addColumn = () => {
                        setHeaders(h => [...h, `Column ${h.length + 1}`]);
                        setRows(r => r.map(row => [...row, ""]));
                    };

                    const addRow = () => {
                        setRows(r => [...r, headers.map(() => "")]);
                    };

                    const removeColumn = (ci: number) => {
                        setHeaders(h => h.filter((_, i) => i !== ci));
                        setRows(r => r.map(row => row.filter((_, i) => i !== ci)));
                    };

                    const removeRow = (ri: number) => {
                        setRows(r => r.filter((_, i) => i !== ri));
                    };

                    const save = async () => {
                        const updated = [...fullContent];
                        updated[index] = {
                            content_type: "table_horizontal",
                            headers,
                            rows,
                        };

                        await fetch(
                            `${config.backendURL}/api/project-report-sections/${sectionId}`,
                            {
                                method: "PUT",
                                credentials: "include",
                                headers: { "Content-Type": "application/json" },
                                body: JSON.stringify({ content: updated }),
                            }
                        );

                        setEditing(false);
                    };

                    const isWideTable = headers.length > 6;

                    return (
                        <Box className="bg-white p-3 rounded-md" sx={{
                            overflowX: 'auto',
                            width: '100%',
                            maxWidth: '100%',
                            position: 'relative'
                        }}>

                            {!editing ? (
                                <>
                                    <table className="text-sm border" style={{
                                        width: isWideTable ? '100%' : 'auto',
                                        tableLayout: isWideTable ? 'fixed' : 'auto',
                                        minWidth: '100%'
                                    }}>
                                        <thead className="bg-gray-50">
                                            <tr>
                                                {headers.map((h, i) => (
                                                    <th key={i} className="border px-3 py-2" style={{
                                                        minWidth: isWideTable ? '120px' : 'auto',
                                                        maxWidth: isWideTable ? '200px' : 'auto',
                                                        wordBreak: 'break-word'
                                                    }}>
                                                        {h}
                                                    </th>
                                                ))}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {rows.map((row, ri) => (
                                                <tr key={ri}>
                                                    {row.map((cell, ci) => (
                                                        <td key={ci} className="border px-3 py-2" style={{
                                                            minWidth: isWideTable ? '120px' : 'auto',
                                                            maxWidth: isWideTable ? '200px' : 'auto',
                                                            wordBreak: 'break-word'
                                                        }}>
                                                            {cell}
                                                        </td>
                                                    ))}
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>

                                    {isWideTable && (
                                        <Chip
                                            icon={<Rotate90DegreesCcw />}
                                            label="Landscape Page"
                                            color="warning"
                                            size="small"
                                            sx={{ mt: 1 }}
                                        />
                                    )}

                                    <Button size="small" onClick={() => setEditing(true)} className="mt-2">
                                        Edit Table
                                    </Button>
                                </>
                            ) : (
                                <>
                                    <Stack direction="row" spacing={1} mb={1}>
                                        <Button size="small" startIcon={<AddOutlined />} onClick={addColumn}>
                                            Add Column
                                        </Button>
                                        <Button size="small" startIcon={<AddOutlined />} onClick={addRow}>
                                            Add Row
                                        </Button>
                                        {isWideTable && (
                                            <Chip
                                                icon={<Rotate90DegreesCcw />}
                                                label="Wide Table (Landscape)"
                                                color="warning"
                                                size="small"
                                            />
                                        )}
                                    </Stack>

                                    <table className="min-w-full text-sm border">
                                        <thead className="bg-gray-50">
                                            <tr>
                                                {headers.map((h, hi) => (
                                                    <th key={hi} className="border px-2 py-1">
                                                        <Stack direction="row" spacing={1}>
                                                            <TextField
                                                                size="small"
                                                                value={h}
                                                                onChange={(e) =>
                                                                    setHeaders(prev =>
                                                                        prev.map((v, i) =>
                                                                            i === hi ? e.target.value : v
                                                                        )
                                                                    )
                                                                }
                                                            />
                                                            <IconButton
                                                                size="small"
                                                                onClick={() => removeColumn(hi)}
                                                            >
                                                                <DeleteOutline fontSize="small" />
                                                            </IconButton>
                                                        </Stack>
                                                    </th>
                                                ))}
                                            </tr>
                                        </thead>
                                        <tbody>
                                            {rows.map((row, ri) => (
                                                <tr key={ri}>
                                                    {row.map((cell, ci) => (
                                                        <td key={ci} className="border px-2 py-1">
                                                            <TextField
                                                                size="small"
                                                                value={cell}
                                                                onChange={(e) =>
                                                                    setRows(prev =>
                                                                        prev.map((r, i) =>
                                                                            i === ri
                                                                                ? r.map((c, j) =>
                                                                                    j === ci ? e.target.value : c
                                                                                )
                                                                                : r
                                                                        )
                                                                    )
                                                                }
                                                            />
                                                        </td>
                                                    ))}
                                                    <td className="border px-2 py-1">
                                                        <IconButton
                                                            size="small"
                                                            onClick={() => removeRow(ri)}
                                                        >
                                                            <DeleteOutline fontSize="small" />
                                                        </IconButton>
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>

                                    <Stack direction="row" spacing={1} mt={2}>
                                        <Button size="small" variant="contained" onClick={save}>
                                            Save
                                        </Button>
                                        <Button size="small" onClick={() => setEditing(false)}>
                                            Cancel
                                        </Button>
                                    </Stack>
                                </>
                            )}
                        </Box>
                    );
                };

                return <TableEditor />;
            }

            if (content.content_type === "table_vertical") {
                const TableEditor = () => {
                    const [editing, setEditing] = useState(false);
                    const [rows, setRows] = useState([...content.content]);

                    const addRow = () => {
                        setRows(r => [...r, { field_name: "", value: "" }]);
                    };

                    const removeRow = (i: number) => {
                        setRows(r => r.filter((_, idx) => idx !== i));
                    };

                    const save = async () => {
                        const updated = [...fullContent];
                        updated[index] = {
                            content_type: "table_vertical",
                            content: rows,
                        };

                        await fetch(
                            `${config.backendURL}/api/project-report-sections/${sectionId}`,
                            {
                                method: "PUT",
                                credentials: "include",
                                headers: { "Content-Type": "application/json" },
                                body: JSON.stringify({ content: updated }),
                            }
                        );

                        setEditing(false);
                    };

                    return (
                        <Box className="bg-white p-3 rounded-md" sx={{ position: 'relative' }}>
                            {!editing ? (
                                <>
                                    <table className="min-w-full text-sm border">
                                        <tbody>
                                            {rows.map((r, i) => (
                                                <tr key={i}>
                                                    <td className="border px-3 py-2 font-medium bg-gray-50" style={{ width: '30%' }}>
                                                        {r.field_name}
                                                    </td>
                                                    <td className="border px-3 py-2" style={{ width: '70%' }}>
                                                        {r.value}
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>

                                    <Button size="small" onClick={() => setEditing(true)} className="mt-2">
                                        Edit Table
                                    </Button>
                                </>
                            ) : (
                                <>
                                    <Button
                                        size="small"
                                        startIcon={<AddOutlined />}
                                        onClick={addRow}
                                        className="mb-2"
                                    >
                                        Add Row
                                    </Button>

                                    <table className="min-w-full text-sm border">
                                        <tbody>
                                            {rows.map((r, i) => (
                                                <tr key={i}>
                                                    <td className="border px-2 py-1" style={{ width: '30%' }}>
                                                        <TextField
                                                            size="small"
                                                            fullWidth
                                                            value={r.field_name}
                                                            onChange={(e) =>
                                                                setRows(prev =>
                                                                    prev.map((row, idx) =>
                                                                        idx === i
                                                                            ? { ...row, field_name: e.target.value }
                                                                            : row
                                                                    )
                                                                )
                                                            }
                                                        />
                                                    </td>
                                                    <td className="border px-2 py-1" style={{ width: '70%' }}>
                                                        <TextField
                                                            size="small"
                                                            fullWidth
                                                            value={r.value}
                                                            onChange={(e) =>
                                                                setRows(prev =>
                                                                    prev.map((row, idx) =>
                                                                        idx === i
                                                                            ? { ...row, value: e.target.value }
                                                                            : row
                                                                    )
                                                                )
                                                            }
                                                        />
                                                    </td>
                                                    <td className="border px-2 py-1" style={{ width: '10%' }}>
                                                        <IconButton
                                                            size="small"
                                                            onClick={() => removeRow(i)}
                                                        >
                                                            <DeleteOutline fontSize="small" />
                                                        </IconButton>
                                                    </td>
                                                </tr>
                                            ))}
                                        </tbody>
                                    </table>

                                    <Stack direction="row" spacing={1} mt={2}>
                                        <Button size="small" variant="contained" onClick={save}>
                                            Save
                                        </Button>
                                        <Button size="small" onClick={() => setEditing(false)}>
                                            Cancel
                                        </Button>
                                    </Stack>
                                </>
                            )}
                        </Box>
                    );
                };

                return <TableEditor />;
            }

            return null;
        },
        []
    );

    // Export function
    const handleExportToDocx = useCallback(async () => {
        const checkedByGroupForExport = groups
            .map(g => ({
                id: g.id,
                title: g.title,
                checked: !!g.checked,
                items: g.items.filter(t => t.checked),
                content: g.content,
            }))
            .filter(g => g.checked || g.items.length > 0);

        await exportToDocx(project, bank, client, checkedByGroupForExport, DEFAULT_IMAGES);
    }, [project, bank, client, groups]);

    // Calculate checked groups
    const checkedByGroup = useMemo(() =>
        groups
            .map(g => ({
                id: g.id,
                main_id: g.main_id,
                title: g.title,
                checked: !!g.checked,
                items: g.items.filter(t => t.checked),
                content: g.content,
            }))
            .filter(g =>
                g.checked ||
                g.items.length > 0 ||
                (Array.isArray(g.content) && g.content.length > 0)
            ),
        [groups]
    );

    const loadStatusLogs = useCallback(async () => {
        if (!project?.id) return;

        try {
            const data = await safeFetch(
                `/api/report-version-status-logs/${project.id}`
            );
            setStatusLogs(Array.isArray(data) ? data : []);
        } catch (e) {
            console.error("Failed to load status logs", e);
        }
    }, [project?.id, safeFetch]);

    useEffect(() => {
        loadStatusLogs();
    }, [loadStatusLogs]);

    useEffect(() => {
        const map: Record<string, boolean> = {};

        processingJobs.forEach((job: ProcessingJobLog) => {
            if (job.status === 'processing') {
                job.tocItems?.forEach((item: TocItem) => {
                    if (item.project_report_sections_id) {
                        map[item.project_report_sections_id] = true;
                    }
                });
            }
        });

        setGeneratingMap(map);
    }, [processingJobs]);

    // Loading state
    const loading = projectLoading || !reportV;

    const prepareSummaryPayload = useCallback(() => {
        if (!project || !reportV) return;

        const sections = groups
            .filter(g => g.checked)
            .map(g => ({
                id: g.id,
                title: g.title,
                isNew: true,
                selected: true,
            }));

        const subsections = groups.flatMap(g =>
            g.items
                .filter(t => t.checked)
                .map(t => {
                    const isNew = isNewSubsection(
                        t.id,
                        reportV.id!,
                        projectRS
                    );

                    return {
                        id: t.id,
                        section_id: g.id,
                        label: t.label,
                        isNew,
                        selected: true,
                    };
                })
        );

        setSummaryPayload({
            project_id: project.id!,
            report_version_id: reportV.id!,

            sections: groups
                .filter(g => g.checked)
                .map(g => ({
                    id: g.id,
                    title: g.title,
                    isNew: false,
                    selected: true,
                })),

            subsections: groups.flatMap(g =>
                g.items
                    .filter(i => i.checked)
                    .map(i => ({
                        id: i.id,
                        label: i.label,
                        isNew: false,
                        section_id: g.id,
                        selected: true,
                    }))
            ),
        });
        setSummaryPopupOpen(true);
    }, [groups, project, reportV, projectRS]);

    const isNewSubsection = (
        subsectionId: string,
        currentVersionId: string,
        projectRS: ProjectReportSectionsP[]
    ) => {
        return !projectRS.some(
            prs =>
                prs.subsection_template_id === subsectionId &&
                prs.report_version_id !== currentVersionId &&
                Array.isArray(prs.content) &&
                prs.content.length > 0
        );
    };

    const toggleAsideMenu = useCallback(() => {
        setIsAsideMenuOpen(prev => !prev);
    }, []);

    const handleBulkQuerySubmit = useCallback(async (query: string) => {
        if (!pendingBulkGeneration || !project || !reportV) return;

        setShowBulkQueryDialog(false);

        const toc = pendingBulkGeneration.sections
            .filter(sec => sec.selected)
            .flatMap(sec => {
                const group = groups.find(g => g.id === sec.id);

                const selectedSubs = pendingBulkGeneration.subsections.filter(
                    sub => sub.section_id === sec.id && sub.selected
                );

                const sectionRow = {
                    section_id: sec.id,
                    section_number: String(group?.order_index ?? ""),
                    section_name: sec.title,

                    subsection_id: "",
                    subsection_number: "",
                    subsection_name: "",
                    query: query || "",

                    project_report_sections_id: group?.main_id ?? null,
                };

                const subsectionRows = selectedSubs.map(sub => ({
                    section_id: sec.id,
                    section_number: String(group?.order_index ?? ""),
                    section_name: sec.title,

                    subsection_id: sub.id,
                    subsection_number: String(
                        group?.items.find(i => i.id === sub.id)?.order_sub_index ?? ""
                    ),
                    subsection_name: sub.label,
                    query: query || "",

                    project_report_sections_id:
                        projectRS.find(p =>
                            p.section_template_id === sub.section_id &&
                            p.subsection_template_id === sub.id &&
                            p.report_version_id === reportV.id
                        )?.id ?? null,
                }));

                return [sectionRow, ...subsectionRows];
            });

        const payload = {
            project_id: project.id!,
            generation_type: Number(reportV.version_number) === 1 ? "new" : "old",
            toc,
        };

        console.log("📤 BULK OL-GPT REQUEST BODY:", payload);

        try {
            const res = await fetch(`${config.backendURL}/api/reports/generate`, {
                method: "POST",
                credentials: "include",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(payload),
            });

            if (!res) return;


            const data = await res.json();

            console.log("RESPONSE BODY:", data);

            setSummaryPopupOpen(false);
        } catch (error) {
            console.error("Failed to generate bulk content:", error);
            setSnackbar({
                open: true,
                message: "Failed to generate content. Please try again.",
                severity: 'error'
            });
        } finally {
            setPendingBulkGeneration(null);
        }
    }, [pendingBulkGeneration, project, reportV, groups, projectRS]);

    return (
        <Box className="min-h-[600px] bg-white p-6 pt-10 relative print:static">


            <GenerationQueryDialog
                open={showBulkQueryDialog}
                onClose={() => {
                    setShowBulkQueryDialog(false);
                    setPendingBulkGeneration(null);
                }}
                onSubmit={handleBulkQuerySubmit}
                title="Add Instructions for Bulk Generation"
                loading={false}
            />

            {/* Review Confirmation Dialog */}
            <Dialog open={!!reviewingSection} onClose={() => setReviewingSection(null)} maxWidth="xs" fullWidth sx={{
                zIndex: 20000
            }}>
                <DialogTitle>Review Section</DialogTitle>
                <DialogContent>
                    <Typography variant="body2" sx={{ mb: 2 }}>
                        Are you sure you want to mark this section as {reviewingSection ? "approved/rejected" : ""}?
                    </Typography>
                    <Typography variant="caption" color="text.secondary">
                        Section: {reviewingSection?.sectionName || "Unknown"}
                    </Typography>
                </DialogContent>
                <DialogActions>
                    <Button onClick={() => setReviewingSection(null)}>Cancel</Button>
                    <Button
                        onClick={() => {
                            if (reviewingSection) {
                                handleReviewSection(reviewingSection.jobId, reviewingSection.projectReportSectionsId, 'rejected');
                            }
                        }}
                        color="error"
                        variant="outlined"
                    >
                        Reject
                    </Button>
                    <Button
                        onClick={() => {
                            if (reviewingSection) {
                                handleReviewSection(reviewingSection.jobId, reviewingSection.projectReportSectionsId, 'approved');
                            }
                        }}
                        color="success"
                        variant="contained"
                    >
                        Approve
                    </Button>
                </DialogActions>
            </Dialog>

            {/* Header Controls */}
            <Box className="mb-2 fixed top-[7vh] left-[200px] right-0 bg-white z-[999] flex items-center gap-3 no-print py-2 px-3">
                <Typography variant="h6" className="font-semibold text-sm">Report</Typography>
                <Chip
                    color="success"
                    label="LIE Report"
                    size="small"
                    sx={{
                        height: 16,
                        '& .MuiChip-label': {
                            fontSize: '0.65rem',
                            paddingInline: '4px',
                        }
                    }}
                />

                {reportV?.status === "Processing" || reportV?.status === "NYS" ? (
                    <Chip
                        color={reportV.status === "NYS" ? "error" : "warning"}
                        label={`Status: ${reportV.status === "NYS" ? "Not Yet Started" : "Processing..."}`}
                        size="small"
                        sx={{
                            height: 16,
                            '& .MuiChip-label': {
                                fontSize: '0.65rem',
                                paddingInline: '5px',
                            }
                        }}
                    />
                ) : (
                    <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
                        <Stack direction="row" spacing={0}>
                            {STAGES.map(s => (
                                <Button
                                    key={s.label}
                                    size="small"
                                    className={`${permissionService.hasPermission(s.key)
                                        ? 'pointer-events-auto'
                                        : 'pointer-events-none opacity-50'
                                        }`}
                                    onClick={() => {
                                        if (permissionService.hasPermission(s.key)) {
                                            setConfirmStage(s.stage);
                                        }
                                    }}
                                    variant="contained"
                                    color={reportV?.status === s.stage ? 'primary' : 'inherit'}
                                    sx={{
                                        textTransform: 'none',
                                        minWidth: 80,
                                        fontSize: '0.5rem',
                                        py: 1,
                                        px: 2,
                                        ml: "-10px",
                                        lineHeight: 1.2,
                                        borderRadius: 0,
                                    }}
                                    style={{
                                        clipPath: "polygon(0 0, calc(100% - 10px) 0,100% 50%, calc(100% - 10px) 100%, 0% 100%, 10px 50%)"
                                    }}
                                    aria-pressed={reportV?.status === s.stage}                                 >
                                    {s.label}
                                </Button>
                            ))}
                        </Stack>
                        <Chip
                            label={STAGES.find(s => s.stage === stage)?.label || stage}
                            size="small"
                            color="default"
                            sx={{ ml: 1, fontSize: '0.6rem' }}
                        />
                    </Box>
                )}
                <Dialog
                    open={!!confirmStage}
                    onClose={() => setConfirmStage(null)}
                    maxWidth="xs"
                    fullWidth
                >
                    <DialogTitle sx={{ fontSize: '1rem', fontWeight: 500 }}>
                        Change Report Stage
                    </DialogTitle>

                    <DialogContent>
                        <Typography sx={{ fontSize: '0.85rem', color: '#374151' }}>
                            Are you sure you want to move this report to
                            <strong> {confirmStage}</strong> stage?
                        </Typography>
                    </DialogContent>

                    <DialogActions>
                        <Button
                            size="small"
                            onClick={() => setConfirmStage(null)}
                        >
                            Cancel
                        </Button>

                        <Button
                            size="small"
                            variant="contained"
                            sx={{ backgroundColor: '#032F5D' }}
                            onClick={() => {
                                if (confirmStage) {
                                    updateReportStatus(confirmStage);
                                }
                                setConfirmStage(null);
                            }}
                        >
                            Confirm
                        </Button>
                    </DialogActions>
                </Dialog>

                <Button
                    className="ms-auto"
                    size="small"
                    variant="contained"
                    endIcon={<AutoAwesome />}
                    sx={{ backgroundColor: "#032F5D", fontSize: '0.6rem' }}
                    onClick={() => {
                        setOpenStatus(true);
                        fetchProcessingJobs();
                    }}
                >
                    Status
                </Button>

                <Button
                    size="small"
                    sx={{ backgroundColor: "#032F5D", paddingInline: "0px", minWidth: "40px" }}
                    variant="contained"
                    className="bg-primary-800"
                    onClick={toggleAsideMenu}
                >
                    <MenuIcon style={{ fontSize: 18 }} />
                </Button>
            </Box>
            <Drawer
                anchor="right"
                open={openStatus}
                onClose={() => setOpenStatus(false)}
                sx={{ zIndex: 9999 }}
            >
                <Box
                    sx={{
                        width: 500,
                        p: 3,
                        backgroundColor: '#ffffff',
                        height: '100%',
                        overflow: 'auto'
                    }}
                >
                    <Box sx={{ mb: 3 }}>
                        <Typography
                            variant="h6"
                            sx={{ fontWeight: 400, color: '#1f2937', fontSize: '1rem' }}
                        >
                            Processing Status
                        </Typography>
                        <Typography
                            variant="caption"
                            sx={{ color: '#6b7280' }}
                        >
                            Live OL-GPT job progress for this project
                        </Typography>
                    </Box>

                    {loadingStatus && (
                        <Box sx={{ py: 4, textAlign: 'center' }}>
                            <CircularProgress size={24} />
                            <Typography sx={{ fontSize: '0.85rem', color: '#6b7280', mt: 1 }}>
                                Fetching job status…
                            </Typography>
                        </Box>
                    )}

                    {!loadingStatus && processingJobs.length === 0 && (
                        <Box
                            sx={{
                                py: 5,
                                textAlign: 'center',
                                borderRadius: 2,
                                backgroundColor: '#f9fafb'
                            }}
                        >
                            <Typography sx={{ fontSize: '0.85rem', color: '#6b7280' }}>
                                No processing jobs at the moment
                            </Typography>
                        </Box>
                    )}

                    <Box sx={{ display: 'flex', flexDirection: 'column', gap: 3 }}>
                        {Array.isArray(processingJobs) && processingJobs.map((job) => (
                            <Box
                                key={job.id}
                                sx={{
                                    p: 2,
                                    borderRadius: 2,
                                    backgroundColor: '#f8fafc',
                                    border: '1px solid #e5e7eb'
                                }}
                            >
                                <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', mb: 2 }}>
                                    <Box>
                                        <Typography
                                            sx={{
                                                fontSize: '0.7rem',
                                                color: '#111827',
                                                mb: 0.5
                                            }}
                                        >
                                            Job ID: {job.olgpt_job_id}
                                        </Typography>
                                        <Typography
                                            sx={{
                                                fontSize: '0.65rem',
                                                color: '#6b7280'
                                            }}
                                        >
                                            Created: {new Date(job.created_at).toLocaleString()}
                                        </Typography>
                                    </Box>
                                    <Chip
                                        size="small"
                                        label={job.status.toUpperCase()}
                                        sx={{
                                            backgroundColor:
                                                job.status === 'processing'
                                                    ? '#FEF3C7'
                                                    : job.status === 'completed'
                                                        ? '#DCFCE7'
                                                        : '#FEE2E2',
                                            color:
                                                job.status === 'processing'
                                                    ? '#92400E'
                                                    : job.status === 'completed'
                                                        ? '#166534'
                                                        : '#991B1B',
                                            fontSize: '0.65rem',
                                            height: 22
                                        }}
                                    />
                                </Box>

                                {job.tocItems && job.tocItems.length > 0 ? (
                                    <Box sx={{ mt: 2 }}>
                                        <Typography variant="subtitle2" sx={{ fontSize: '0.75rem', fontWeight: 600, mb: 1, color: '#374151' }}>
                                            Generated Sections ({job.tocItems.length})
                                        </Typography>
                                        <Box sx={{ display: 'flex', flexDirection: 'column', gap: 1.5 }}>
                                            {job.tocItems.map((item: TocItem, index: number) => (
                                                <Paper
                                                    key={index}
                                                    variant="outlined"
                                                    sx={{
                                                        p: 1.5,
                                                        backgroundColor: item.status === 'approved'
                                                            ? '#f0fdf4'
                                                            : item.status === 'rejected'
                                                                ? '#fef2f2'
                                                                : '#f8fafc'
                                                    }}
                                                >
                                                    <Box sx={{ display: 'flex', justifyContent: 'space-between', alignItems: 'flex-start', mb: 1 }}>
                                                        <Box>
                                                            <Typography sx={{ fontSize: '0.75rem', fontWeight: 500 }}>
                                                                {item.section_name} {item.subsection_name && `> ${item.subsection_name}`}
                                                            </Typography>
                                                            {item.query && (
                                                                <Typography sx={{ fontSize: '0.65rem', color: '#6b7280', mt: 0.5 }}>
                                                                    Query: "{item.query}"
                                                                </Typography>
                                                            )}
                                                        </Box>
                                                        <Box sx={{ display: 'flex', gap: 0.5 }}>
                                                            {item.status === 'pending' && (
                                                                <>
                                                                    <IconButton
                                                                        size="small"
                                                                        onClick={() => setReviewingSection({
                                                                            jobId: job.olgpt_job_id,
                                                                            projectReportSectionsId: item.project_report_sections_id,
                                                                            sectionName: `${item.section_name} ${item.subsection_name ? `> ${item.subsection_name}` : ''}`
                                                                        })}
                                                                        sx={{ color: '#10b981' }}
                                                                        title="Approve"
                                                                    >
                                                                        <ThumbUp fontSize="small" />
                                                                    </IconButton>
                                                                    <IconButton
                                                                        size="small"
                                                                        onClick={() => setReviewingSection({
                                                                            jobId: job.olgpt_job_id,
                                                                            projectReportSectionsId: item.project_report_sections_id,
                                                                            sectionName: `${item.section_name} ${item.subsection_name ? `> ${item.subsection_name}` : ''}`
                                                                        })}
                                                                        sx={{ color: '#ef4444' }}
                                                                        title="Reject"
                                                                    >
                                                                        <ThumbDown fontSize="small" />
                                                                    </IconButton>
                                                                </>
                                                            )}
                                                            {item.status === 'approved' && (
                                                                <Chip
                                                                    size="small"
                                                                    icon={<CheckCircle fontSize="small" />}
                                                                    label="Approved"
                                                                    color="success"
                                                                    sx={{ height: 22, fontSize: '0.6rem' }}
                                                                />
                                                            )}
                                                            {item.status === 'rejected' && (
                                                                <Chip
                                                                    size="small"
                                                                    icon={<Cancel fontSize="small" />}
                                                                    label="Rejected"
                                                                    color="error"
                                                                    sx={{ height: 22, fontSize: '0.6rem' }}
                                                                />
                                                            )}
                                                        </Box>
                                                    </Box>
                                                    {item.content_blocks && item.content_blocks.length > 0 && (
                                                        <Box
                                                            sx={{
                                                                mt: 1,
                                                                display: 'flex',
                                                                justifyContent: 'space-between',
                                                                alignItems: 'center'
                                                            }}
                                                        >
                                                            <Typography sx={{ fontSize: '0.7rem', color: '#4b5563' }}>
                                                                Preview: {item.content_blocks[0]?.value?.substring(0, 60)}...
                                                            </Typography>

                                                            <Button
                                                                size="small"
                                                                variant="text"
                                                                sx={{ fontSize: '0.65rem' }}
                                                                onClick={() => {
                                                                    setPreviewContent(item.content_blocks[0]?.value);
                                                                    // setReviewingSection({
                                                                    //     jobId: job.id,
                                                                    //     subsectionId: item.subsection_id || item.section_id,
                                                                    // });
                                                                    setOpenPreview(true);
                                                                }}
                                                            >
                                                                Options
                                                            </Button>
                                                        </Box>
                                                    )}
                                                </Paper>
                                            ))}
                                        </Box>
                                    </Box>
                                ) : (
                                    <Typography sx={{ fontSize: '0.75rem', color: '#6b7280', textAlign: 'center', py: 2 }}>
                                        No sections generated yet
                                    </Typography>
                                )}
                            </Box>
                        ))}
                    </Box>
                    <Dialog
                        open={openPreview}
                        onClose={() => setOpenPreview(false)}
                        maxWidth="md"
                        fullWidth
                        sx={{
                            zIndex: 20000,
                            '& .MuiDialog-paper': {
                                borderRadius: 3,
                                minHeight: '60vh',
                            }
                        }}
                    >
                        {/* Header */}
                        <DialogTitle
                            sx={{
                                px: 3,
                                py: 2,
                                display: 'flex',
                                alignItems: 'center',
                                justifyContent: 'space-between',
                                borderBottom: '1px solid #E5E7EB'
                            }}
                        >
                            <Typography sx={{ fontSize: '1rem', fontWeight: 600 }}>
                                Generated Content
                            </Typography>

                            <Button
                                size="small"
                                variant="outlined"
                                sx={{ textTransform: 'none' }}
                            >
                                Edit
                            </Button>
                        </DialogTitle>

                        {/* Content */}
                        <DialogContent
                            sx={{
                                px: 3,
                                py: 2,
                                backgroundColor: '#F9FAFB'
                            }}
                        >
                            <Typography
                                sx={{
                                    fontSize: '0.9rem',
                                    lineHeight: 1.7,
                                    whiteSpace: 'pre-wrap',
                                    color: '#374151'
                                }}
                            >
                                {previewContent}
                            </Typography>
                        </DialogContent>

                        {/* Actions */}
                        <DialogActions
                            sx={{
                                px: 3,
                                py: 2,
                                borderTop: '1px solid #E5E7EB',
                                flexDirection: 'column',
                                alignItems: 'stretch',
                                gap: 1.5
                            }}
                        >
                            <Box sx={{ display: 'flex', justifyContent: 'between', gap: 1 }}>
                                <TextField
                                    fullWidth
                                    label="Improvement instructions"
                                    placeholder="Make it more technical, remove references, add Indian standards"
                                    size="small"
                                />
                                <Box sx={{ display: 'flex', justifyContent: 'between', gap: 1 }}>
                                    <Button
                                        variant="contained"
                                        sx={{ px: 1.5 }}
                                    // onClick={handleRegenerate}
                                    >
                                        ReGenerate
                                    </Button>
                                    <Button
                                        onClick={() => setOpenPreview(false)}
                                        variant="text"
                                    >
                                        Close
                                    </Button>

                                </Box>
                            </Box>
                        </DialogActions>
                    </Dialog>
                </Box>
            </Drawer>


            <Box className="flex gap-2">
                {/* Main Content */}
                <Box className="grid grid-cols-12 w-[80%]">
                    {/* Sidebar */}
                    <Box
                        component="aside"
                        className={`col-span-12 fixed z-[99] top-0 right-0 lg:col-span-3 min-w-[350px] shadow-md transition-transform duration-300 ease-in-out`}
                        style={{ transform: `translateX(${isAsideMenuOpen ? "0" : "100%"})` }}
                    >
                        <Paper variant="outlined" className="p-2 max-h-screen h-max no-print">
                            <DndContext
                                sensors={sensors}
                                collisionDetection={closestCenter}
                                onDragStart={handleDragStart}
                                onDragOver={handleDragOver}
                                onDragEnd={handleDragEnd}
                            >
                                <div className="pt-[16vh] h-screen flex flex-col">
                                    {/* Scrollable content */}
                                    <div className="flex-1 overflow-y-auto">
                                        <SortableContext
                                            items={groups.map(g => g.id)}
                                            strategy={verticalListSortingStrategy}
                                        >
                                            {groups
                                                .sort((a, b) => Number(a.order_index) - Number(b.order_index))
                                                .map(g => (
                                                    <SortableGroup
                                                        key={g.id}
                                                        group={g}
                                                        setCurrentCheck={setCurrentCheck}
                                                        onToggle={toggleGroup}
                                                        onToggleCheck={toggleGroupAllChecked}
                                                        onToggleTask={toggleTaskChecked}
                                                        allSection={allSection}
                                                        reportV={reportV}
                                                        safeFetch={safeFetch}
                                                        setAllSection={setAllSection}
                                                        setAllSubSection={setAllSubSection}
                                                        projectRS={projectRS}
                                                        project={project}
                                                        onGenerateContent={handleGenerateContent}
                                                        generatingMap={generatingMap}
                                                        setSnackbar={setSnackbar}
                                                    />
                                                ))}
                                        </SortableContext>

                                        {/* Add Section Input */}
                                        <div className="add-section-btn mt-4 p-3 border-t border-gray-300 grid grid-cols-[1fr_40px] gap-2 items-center">
                                            <input
                                                type="text"
                                                placeholder="Enter section name"
                                                className="w-full border border-gray-300 rounded px-2 py-1 text-sm focus:outline-none focus:ring-1 focus:ring-blue-500"
                                                value={newSectionName}
                                                onChange={(e) => setNewSectionName(e.target.value)}
                                                onKeyPress={(e) => {
                                                    if (e.key === "Enter") handleAddSection();
                                                }}
                                            />
                                            <button
                                                className="w-full h-full py-2 text-sm bg-blue-600 text-white rounded hover:bg-blue-700 disabled:bg-gray-400"
                                                disabled={!newSectionName.trim()}
                                                onClick={handleAddSection}
                                            >
                                                +
                                            </button>
                                        </div>
                                    </div>

                                    {/* Fixed bottom button */}
                                    <div className="shrink-0 border-t bg-white py-3 flex justify-center">
                                        <Button
                                            className="px-2 py-1 text-sm"
                                            size="small"
                                            startIcon={<AutoAwesome />}
                                            sx={{ backgroundColor: "#032F5D" }}
                                            onClick={prepareSummaryPayload}
                                            variant="contained"
                                            color="primary"
                                        >
                                            Generate Summary
                                        </Button>
                                    </div>
                                </div>
                            </DndContext>
                        </Paper>
                    </Box>

                    <Dialog open={summaryPopupOpen} onClose={() => setSummaryPopupOpen(false)} maxWidth="sm" fullWidth>
                        <DialogTitle>Confirm Summary Generation</DialogTitle>

                        <DialogContent dividers>
                            {summaryPayload?.sections.map(section => {
                                const sectionSubs = summaryPayload.subsections.filter(
                                    sub => sub.section_id === section.id
                                );

                                return (
                                    <div key={section.id} className="mb-3">
                                        {/* Section checkbox */}
                                        <label className="flex items-center gap-2 font-medium text-sm">
                                            <input
                                                type="checkbox"
                                                checked={section.selected}
                                                onChange={() =>
                                                    setSummaryPayload(prev =>
                                                        prev && {
                                                            ...prev,
                                                            sections: prev.sections.map(s =>
                                                                s.id === section.id
                                                                    ? { ...s, selected: !s.selected }
                                                                    : s
                                                            ),
                                                            subsections: prev.subsections.map(sub =>
                                                                sub.section_id === section.id
                                                                    ? { ...sub, selected: !section.selected }
                                                                    : sub
                                                            ),
                                                        }
                                                    )
                                                }
                                            />
                                            {section.title}
                                        </label>

                                        {/* Subsections */}
                                        <div className="ml-6 mt-1 space-y-1">
                                            {sectionSubs.map(sub => (
                                                <label
                                                    key={sub.id}
                                                    className="flex items-center gap-2 text-sm"
                                                >
                                                    <input
                                                        type="checkbox"
                                                        checked={sub.selected}
                                                        onChange={() =>
                                                            setSummaryPayload(prev =>
                                                                prev && {
                                                                    ...prev,
                                                                    subsections: prev.subsections.map(s =>
                                                                        s.id === sub.id
                                                                            ? { ...s, selected: !s.selected }
                                                                            : s
                                                                    ),
                                                                }
                                                            )
                                                        }
                                                    />
                                                    {sub.label}
                                                    {sub.isNew && (
                                                        <span className="text-xs text-green-600 ml-1">(new)</span>
                                                    )}
                                                </label>
                                            ))}
                                        </div>
                                    </div>
                                );
                            })}

                        </DialogContent>

                        <DialogActions>
                            <Button onClick={() => setSummaryPopupOpen(false)}>Cancel</Button>
                            <Button
                                variant="contained"
                                onClick={async () => {
                                    if (!summaryPayload || !project || !reportV) return;

                                    setPendingBulkGeneration({
                                        sections: summaryPayload.sections.filter(sec => sec.selected),
                                        subsections: summaryPayload.subsections.filter(sub => sub.selected)
                                    });

                                    setSummaryPopupOpen(false);
                                    setShowBulkQueryDialog(true);
                                }}
                            >
                                Generate
                            </Button>
                        </DialogActions>
                    </Dialog>

                    {/* Report Viewer */}
                    <Box className="col-span-12">
                        <ReportViewer
                            project={project}
                            bank={bank}
                            client={client}
                            checkedByGroup={checkedByGroup}
                            renderOlResult={renderOlResult}
                            onExportDocx={handleExportToDocx}
                            onExportPdf={print}
                        />
                    </Box>
                </Box>

                {/* Loading Overlay */}
                {loading && (
                    <Box className="bg-white/75 print:hidden fixed top-0 left-0 w-full h-full flex items-center justify-center z-[9999]">
                        <div className='w-14 h-14' style={{
                            backgroundImage: 'url(https://api.iconify.design/eos-icons:bubble-loading.svg?color=%23032F5D)',
                            backgroundSize: 'contain'
                        }}></div>
                    </Box>
                )}

                {/* Right Side Status Progress */}
                <Box className="w-[20%] h-[100vh] sticky rounded-md p-1 overflow-y-auto no-print">
                    <ReportVersionProgress logs={statusLogs} />
                </Box>
            </Box>
            <Snackbar
                open={snackbar.open}
                autoHideDuration={4000}
                onClose={() => setSnackbar(prev => ({ ...prev, open: false }))}
                anchorOrigin={{ vertical: 'bottom', horizontal: 'left' }}
            >
                <Alert
                    severity={snackbar.severity}
                    onClose={() => setSnackbar(prev => ({ ...prev, open: false }))}
                    sx={{ width: '100%' }}
                >
                    {snackbar.message}
                </Alert>
            </Snackbar>
        </Box>
    );
}