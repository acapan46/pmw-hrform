import {
  AppBar,
  Box,
  Button,
  IconButton,
  Menu,
  MenuItem,
  Stack,
  Toolbar,
  Typography,
  useTheme,
  useMediaQuery,
} from "@mui/material";
import {
  Person as PersonIcon,
  Logout as LogoutIcon,
  Settings as SettingsIcon,
} from "@mui/icons-material";
import { useState } from "react";
import RoleBadge from "./RoleBadge";
import Logo from "../Logo";

interface HeaderProps {
  userEmail: string;
  isAdmin: boolean;
  onLogout: () => void;
  onSwitch: () => void;
  onOpenBuilder?: () => void;
}

export default function Header({
  userEmail,
  isAdmin,
  onLogout,
  onSwitch,
  onOpenBuilder,
}: HeaderProps) {
  const theme = useTheme();
  const isMobile = useMediaQuery(theme.breakpoints.down("sm"));
  const [anchorEl, setAnchorEl] = useState<null | HTMLElement>(null);
  const open = Boolean(anchorEl);

  const handleMenuOpen = (event: React.MouseEvent<HTMLElement>) => {
    setAnchorEl(event.currentTarget);
  };

  const handleMenuClose = () => {
    setAnchorEl(null);
  };

  return (
    <AppBar
      position="sticky"
      elevation={0}
      sx={{
        backgroundColor: "rgba(255, 255, 255, 0.92)",
        backdropFilter: "blur(20px)",
        borderBottom: "1px solid rgba(0, 0, 0, 0.06)",
        boxShadow: "0 1px 3px rgba(0, 0, 0, 0.04)",
        zIndex: theme.zIndex.drawer + 1,
        minHeight: isMobile ? 56 : 64,
      }}
    >
      <Toolbar sx={{ gap: 2, minHeight: "inherit" }}>
        {/* Brand mark */}
        <Box
          sx={{
            display: "flex",
            alignItems: "center",
            gap: 1.5,
          }}
        >
          <Logo size={isMobile ? 32 : 40} />
          <Stack direction="column" spacing={0}>
            <Typography
              variant="h5"
              component="h1"
              sx={{
                fontWeight: 700,
                color: "#111827",
                letterSpacing: "-0.02em",
                lineHeight: 1.2,
                fontSize: "1.25rem",
              }}
            >
              PMW HR
            </Typography>
            <Typography
              variant="caption"
              sx={{
                color: "#6B7280",
                textTransform: "uppercase",
                letterSpacing: "0.1em",
                fontSize: "0.7rem",
                fontWeight: 600,
                lineHeight: 1,
              }}
            >
              Forms Portal
            </Typography>
          </Stack>
        </Box>

        <Box sx={{ flexGrow: 1 }} />

        {isAdmin && (
          <>
            {onOpenBuilder && (
              <Button
                variant="contained"
                startIcon={<SettingsIcon />}
                onClick={onOpenBuilder}
                sx={{
                  mr: 1,
                  borderRadius: "12px",
                  textTransform: "none",
                  backgroundColor: "#6264A7",
                  color: "#ffffff",
                  fontWeight: 600,
                  fontSize: "0.85rem",
                  py: 1,
                  px: 2.5,
                  boxShadow: "0 2px 8px rgba(98, 100, 167, 0.25)",
                  transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
                  "&:hover": {
                    backgroundColor: "#4A4C80",
                    boxShadow: "0 4px 12px rgba(98, 100, 167, 0.35)",
                    transform: "translateY(-1px)",
                  },
                  "&:active": {
                    transform: "scale(0.98) translateY(0)",
                  },
                }}
              >
                Form Builder
              </Button>
            )}

          </>
        )}

        <RoleBadge isAdmin={isAdmin} />

        <IconButton
          onClick={handleMenuOpen}
          size="small"
          sx={{
            ml: 0.5,
            p: 0.75,
            borderRadius: "12px",
            backgroundColor: "rgba(0, 120, 212, 0.06)",
            border: "1px solid rgba(0, 120, 212, 0.1)",
            transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
            "&:hover": {
              backgroundColor: "rgba(0, 120, 212, 0.12)",
              borderColor: "rgba(0, 120, 212, 0.2)",
              transform: "translateY(-1px)",
            },
          }}
        >
          <Box
            sx={{
              width: 32,
              height: 32,
              borderRadius: "10px",
              backgroundColor: "rgba(0, 120, 212, 0.1)",
              display: "flex",
              alignItems: "center",
              justifyContent: "center",
            }}
          >
            <PersonIcon sx={{ fontSize: 18, color: "#0078D4" }} />
          </Box>
        </IconButton>

        <Menu
          anchorEl={anchorEl}
          open={open}
          onClose={handleMenuClose}
          slotProps={{
            paper: {
              sx: {
                minWidth: 220,
                borderRadius: "14px",
                boxShadow: "0 8px 32px rgba(0, 0, 0, 0.12)",
                border: "1px solid rgba(0, 0, 0, 0.06)",
                mt: 1,
              },
            },
          }}
          transformOrigin={{ horizontal: "right", vertical: "top" }}
          anchorOrigin={{ horizontal: "right", vertical: "bottom" }}
        >
          <MenuItem disabled sx={{ cursor: "default", px: 2.5, py: 1.5 }}>
            <Typography
              variant="body2"
              sx={{
                color: "#111827",
                fontWeight: 500,
                maxWidth: 200,
                overflow: "hidden",
                textOverflow: "ellipsis",
                whiteSpace: "nowrap",
              }}
            >
              {userEmail}
            </Typography>
          </MenuItem>
          <MenuItem
            onClick={() => {
              handleMenuClose();
              onSwitch();
            }}
            sx={{ py: 1.25, px: 2.5 }}
          >
            <PersonIcon sx={{ mr: 1.5, fontSize: 20, color: "#6B7280" }} />
            <Typography variant="body2">Switch account</Typography>
          </MenuItem>
          <MenuItem
            onClick={() => {
              handleMenuClose();
              onLogout();
            }}
            sx={{ py: 1.25, px: 2.5 }}
          >
            <LogoutIcon sx={{ mr: 1.5, fontSize: 20, color: "#DC2626" }} />
            <Typography variant="body2" sx={{ color: "#DC2626" }}>
              Sign out
            </Typography>
          </MenuItem>
        </Menu>
      </Toolbar>
    </AppBar>
  );
}