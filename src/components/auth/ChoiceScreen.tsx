import {
  Box,
  Button,
  Card,
  CardContent,
  Checkbox,
  Container,
  Divider,
  FormControlLabel,
  Stack,
  Typography,
  useMediaQuery,
  useTheme,
} from "@mui/material";
import { Login as LoginIcon, Person as PersonIcon } from "@mui/icons-material";
import { fadeInUp } from "../../theme";
import Logo from "../../components/Logo";

interface ChoiceScreenProps {
  onLogin: () => void;
  onGuest: () => void;
}

export default function ChoiceScreen({ onLogin, onGuest }: ChoiceScreenProps) {
  const theme = useTheme();
  const isMobile = useMediaQuery(theme.breakpoints.down("sm"));
  const isTablet = useMediaQuery(theme.breakpoints.between("sm", "md"));

  return (
    <Box
      sx={{
        minHeight: "100vh",
        display: "flex",
        flexDirection: "column",
        justifyContent: "center",
        alignItems: "center",
        background: "#F8F9FC",
        position: "relative",
        overflow: "hidden",
        padding: isMobile ? 2 : isTablet ? 3 : 4,
      }}
    >
      {/* Decorative background - refined gradient mesh */}
      <Box
        sx={{
          position: "absolute",
          top: "-20%",
          right: "-15%",
          width: isMobile ? "350px" : "600px",
          height: isMobile ? "350px" : "600px",
          borderRadius: "50%",
          background: "radial-gradient(circle, rgba(0, 120, 212, 0.05) 0%, rgba(0, 120, 212, 0) 70%)",
          pointerEvents: "none",
          zIndex: 0,
        }}
      />
      <Box
        sx={{
          position: "absolute",
          bottom: "-25%",
          left: "-20%",
          width: isMobile ? "400px" : "700px",
          height: isMobile ? "400px" : "700px",
          borderRadius: "50%",
          background: "radial-gradient(circle, rgba(98, 100, 167, 0.04) 0%, rgba(98, 100, 167, 0) 70%)",
          pointerEvents: "none",
          zIndex: 0,
        }}
      />
      <Box
        sx={{
          position: "absolute",
          top: "35%",
          left: "55%",
          width: isMobile ? "250px" : "450px",
          height: isMobile ? "250px" : "450px",
          borderRadius: "50%",
          background: "radial-gradient(circle, rgba(98, 100, 167, 0.03) 0%, rgba(98, 100, 167, 0) 70%)",
          pointerEvents: "none",
          zIndex: 0,
        }}
      />

      {/* Subtle geometric accent lines */}
      <svg
        style={{
          position: "absolute",
          top: 0,
          left: 0,
          width: "100%",
          height: "100%",
          pointerEvents: "none",
          zIndex: 0,
        }}
        viewBox="0 0 1440 900"
        preserveAspectRatio="xMidYMid slice"
        fill="none"
      >
        <path d="M0 120 Q360 80 720 140 T1440 100" stroke="rgba(0, 120, 212, 0.04)" strokeWidth="1.5" />
        <path d="M0 780 Q360 820 720 760 T1440 800" stroke="rgba(98, 100, 167, 0.04)" strokeWidth="1.5" />
        <circle cx="1200" cy="150" r="100" stroke="rgba(0, 120, 212, 0.03)" strokeWidth="1" />
        <circle cx="180" cy="720" r="70" stroke="rgba(98, 100, 167, 0.03)" strokeWidth="1" />
      </svg>

      <Container maxWidth="sm" sx={{ position: "relative", zIndex: 1 }}>
        <Card
          sx={{
            width: "100%",
            maxWidth: isMobile ? "100%" : 480,
            mx: "auto",
            borderRadius: "24px",
            boxShadow: "0 4px 24px rgba(0, 0, 0, 0.06), 0 1px 3px rgba(0, 0, 0, 0.04)",
            border: "1px solid rgba(0, 0, 0, 0.04)",
            backgroundColor: "rgba(255, 255, 255, 0.92)",
            backdropFilter: "blur(12px)",
            transition: "all 0.3s cubic-bezier(0.4, 0, 0.2, 1)",
            "&:hover": {
              boxShadow: "0 8px 40px rgba(0, 0, 0, 0.1), 0 2px 8px rgba(0, 0, 0, 0.06)",
              transform: "translateY(-2px)",
            },
            animation: `${fadeInUp} 0.6s cubic-bezier(0.16, 1, 0.3, 1) forwards`,
          }}
        >
          {/* Top accent bar */}
          <Box
            sx={{
              height: 4,
              background: "linear-gradient(90deg, #0078D4 0%, #6264A7 100%)",
              borderRadius: "24px 24px 0 0",
            }}
          />

          <CardContent
            sx={{
              padding: isMobile ? 3.5 : 5,
              display: "flex",
              flexDirection: "column",
              alignItems: "center",
            }}
          >
            {/* Logo with subtle glow */}
            <Box
              sx={{
                position: "relative",
                mb: 3,
                "&::before": {
                  content: '""',
                  position: "absolute",
                  top: "50%",
                  left: "50%",
                  transform: "translate(-50%, -50%)",
                  width: isMobile ? 88 : 104,
                  height: isMobile ? 88 : 104,
                  borderRadius: "50%",
                  background: "radial-gradient(circle, rgba(0, 120, 212, 0.15) 0%, rgba(0, 120, 212, 0) 70%)",
                  zIndex: -1,
                },
              }}
            >
              <Logo size={isMobile ? 72 : 88} />
            </Box>

            <Typography
              variant="h2"
              component="h1"
              gutterBottom
              align="center"
              sx={{
                fontWeight: 700,
                color: "#111827",
                letterSpacing: "-0.03em",
                fontSize: isMobile ? "1.75rem" : "2.25rem",
              }}
            >
              PMW HR Forms
            </Typography>

            <Typography
              variant="body1"
              align="center"
              sx={{
                mb: 4,
                maxWidth: 380,
                lineHeight: 1.6,
                color: "#6B7280",
                fontSize: "0.95rem",
              }}
            >
              Sign in with your Microsoft 365 account to access your submission history and track approval status.
            </Typography>

            <Stack spacing={2.5} sx={{ width: "100%" }}>
              <Button
                variant="contained"
                fullWidth
                size="large"
                startIcon={<LoginIcon />}
                onClick={() => {
                  onLogin();
                }}
                sx={{
                  backgroundColor: "#0078D4",
                  color: "#ffffff",
                  borderRadius: "12px",
                  py: 1.75,
                  fontSize: "1rem",
                  fontWeight: 500,
                  boxShadow: "0 2px 8px rgba(0, 120, 212, 0.2)",
                  transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
                  "&:hover": {
                    backgroundColor: "#0068C4",
                    boxShadow: "0 6px 20px rgba(0, 120, 212, 0.3)",
                    transform: "translateY(-1px)",
                  },
                  "&:active": {
                    transform: "scale(0.98) translateY(0)",
                  },
                }}
              >
                Sign in with Microsoft 365
              </Button>

              <FormControlLabel
                control={
                  <Checkbox
                    size="small"
                    sx={{
                      color: "rgba(17, 24, 39, 0.3)",
                      "&.Mui-checked": {
                        color: "#0078D4",
                      },
                    }}
                  />
                }
                label={
                  <Typography variant="body2" sx={{ color: "#6B7280", fontSize: "0.85rem" }}>
                    Remember my choice on this device
                  </Typography>
                }
                sx={{ alignSelf: "flex-start", ml: 0.5 }}
              />

              <Divider>
                <Typography variant="body2" sx={{ color: "#9CA3AF", fontSize: "0.75rem" }}>
                  or
                </Typography>
              </Divider>

              <Button
                variant="outlined"
                fullWidth
                size="large"
                startIcon={<PersonIcon />}
                onClick={onGuest}
                sx={{
                  borderColor: "rgba(17, 24, 39, 0.15)",
                  color: "#6B7280",
                  borderRadius: "12px",
                  py: 1.75,
                  fontSize: "1rem",
                  fontWeight: 500,
                  borderWidth: "1.5px",
                  transition: "all 0.25s cubic-bezier(0.4, 0, 0.2, 1)",
                  "&:hover": {
                    borderColor: "rgba(0, 120, 212, 0.5)",
                    backgroundColor: "rgba(0, 120, 212, 0.04)",
                    color: "#0078D4",
                    transform: "translateY(-1px)",
                  },
                  "&:active": {
                    transform: "scale(0.98) translateY(0)",
                  },
                }}
              >
                Continue as Guest
              </Button>
            </Stack>

            <Typography
              variant="caption"
              align="center"
              sx={{
                mt: 4,
                color: "#9CA3AF",
                fontSize: "0.75rem",
                lineHeight: 1.6,
                maxWidth: 380,
              }}
            >
              Only PMW internal M365 accounts are permitted. Guests can browse publicly available forms.
            </Typography>
          </CardContent>
        </Card>
      </Container>
    </Box>
  );
}