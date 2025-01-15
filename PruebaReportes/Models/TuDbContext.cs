using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace PruebaReportes.Models;

public partial class TuDbContext : DbContext
{
    public TuDbContext()
    {
    }

    public TuDbContext(DbContextOptions<TuDbContext> options)
        : base(options)
    {
    }

    public virtual DbSet<TfaCategory> TfaCategories { get; set; }

    public virtual DbSet<TfaCertificate> TfaCertificates { get; set; }

    public virtual DbSet<TfaHistory> TfaHistories { get; set; }

    public virtual DbSet<TfaRol> TfaRols { get; set; }

    public virtual DbSet<TfaTask> TfaTasks { get; set; }

    public virtual DbSet<TfaTeam> TfaTeams { get; set; }

    public virtual DbSet<TfaTeamstatus> TfaTeamstatuses { get; set; }

    public virtual DbSet<TfaUser> TfaUsers { get; set; }

    public virtual DbSet<TfaUserPoint> TfaUserPoints { get; set; }

    public virtual DbSet<TfaUsersTask> TfaUsersTasks { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        => optionsBuilder.UseSqlServer("Name=DefaultConnection");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.UseCollation("Modern_Spanish_CI_AI");

        modelBuilder.Entity<TfaCategory>(entity =>
        {
            entity.HasKey(e => e.CategoryId).HasName("PK__TFA_CATE__23CAF1F8D2B0FF74");

            entity.ToTable("TFA_CATEGORIES");

            entity.Property(e => e.CategoryId).HasColumnName("categoryID");
            entity.Property(e => e.CategoryDeadLine).HasColumnName("categoryDeadLine");
            entity.Property(e => e.CategoryDescription)
                .HasMaxLength(100)
                .HasColumnName("categoryDescription");
            entity.Property(e => e.CategoryName)
                .HasMaxLength(50)
                .HasColumnName("categoryName");
            entity.Property(e => e.CategoryPoints).HasColumnName("categoryPoints");
        });

        modelBuilder.Entity<TfaCertificate>(entity =>
        {
            entity.HasKey(e => e.CertificatesId).HasName("PK__TFA_CERT__9F43BE1A33A49510");

            entity.ToTable("TFA_CERTIFICATES");

            entity.Property(e => e.CertificatesId).HasColumnName("certificatesID");
            entity.Property(e => e.CertificateCategory).HasColumnName("certificateCategory");
            entity.Property(e => e.CertificateLead).HasColumnName("certificateLead");
            entity.Property(e => e.CertificateMessage)
                .HasMaxLength(100)
                .HasColumnName("certificateMessage");
            entity.Property(e => e.CertificateReceptor).HasColumnName("certificateReceptor");

            entity.HasOne(d => d.CertificateCategoryNavigation).WithMany(p => p.TfaCertificates)
                .HasForeignKey(d => d.CertificateCategory)
                .HasConstraintName("FK_category");

            entity.HasOne(d => d.CertificateLeadNavigation).WithMany(p => p.TfaCertificateCertificateLeadNavigations)
                .HasForeignKey(d => d.CertificateLead)
                .HasConstraintName("FK_certificateLead");

            entity.HasOne(d => d.CertificateReceptorNavigation).WithMany(p => p.TfaCertificateCertificateReceptorNavigations)
                .HasForeignKey(d => d.CertificateReceptor)
                .HasConstraintName("FK_certificateReceptor");
        });

        modelBuilder.Entity<TfaHistory>(entity =>
        {
            entity.HasKey(e => e.HistoryId).HasName("PK__TFA_HIST__19BDBDB38DE91EFE");

            entity.ToTable("TFA_HISTORY");

            entity.Property(e => e.HistoryId).HasColumnName("historyID");
            entity.Property(e => e.HistoryEmission).HasColumnName("historyEmission");
            entity.Property(e => e.ReportType)
                .HasMaxLength(50)
                .IsUnicode(false)
                .HasColumnName("reportType");
            entity.Property(e => e.UserCategoriesId).HasColumnName("userCategoriesID");
            entity.Property(e => e.UserCertificateId).HasColumnName("userCertificateID");
            entity.Property(e => e.UserHistoryId).HasColumnName("userHistoryID");

            entity.HasOne(d => d.UserCategories).WithMany(p => p.TfaHistories)
                .HasForeignKey(d => d.UserCategoriesId)
                .HasConstraintName("FK_userCategoriesID");

            entity.HasOne(d => d.UserCertificate).WithMany(p => p.TfaHistories)
                .HasForeignKey(d => d.UserCertificateId)
                .HasConstraintName("FK_userCertificateID");

            entity.HasOne(d => d.UserHistory).WithMany(p => p.TfaHistories)
                .HasForeignKey(d => d.UserHistoryId)
                .HasConstraintName("FK_userHistoryID");
        });

        modelBuilder.Entity<TfaRol>(entity =>
        {
            entity.HasKey(e => e.RolId).HasName("PK__TFA_ROL__5402365444BEE72A");

            entity.ToTable("TFA_ROL");

            entity.Property(e => e.RolId).HasColumnName("rolID");
            entity.Property(e => e.RolDescription)
                .HasMaxLength(500)
                .IsUnicode(false)
                .HasColumnName("rolDescription");
            entity.Property(e => e.RolName)
                .HasMaxLength(50)
                .HasColumnName("rolName");
        });

        modelBuilder.Entity<TfaTask>(entity =>
        {
            entity.HasKey(e => e.TaskId).HasName("PK__TFA_TASK__DD5D55A251682F04");

            entity.ToTable("TFA_TASKS");

            entity.Property(e => e.TaskId).HasColumnName("taskID");
            entity.Property(e => e.CategoryId).HasColumnName("categoryID");
            entity.Property(e => e.TaskName)
                .HasMaxLength(50)
                .HasColumnName("taskName");

            entity.HasOne(d => d.Category).WithMany(p => p.TfaTasks)
                .HasForeignKey(d => d.CategoryId)
                .HasConstraintName("FK_categoryID");
        });

        modelBuilder.Entity<TfaTeam>(entity =>
        {
            entity.HasKey(e => e.TeamId).HasName("PK__TFA_TEAM__5ED7534A6BB50E22");

            entity.ToTable("TFA_TEAMS");

            entity.Property(e => e.TeamId).HasColumnName("teamID");
            entity.Property(e => e.CategoriesId).HasColumnName("categoriesID");
            entity.Property(e => e.TeamDescription)
                .HasMaxLength(50)
                .HasColumnName("teamDescription");
            entity.Property(e => e.TeamLeadId).HasColumnName("teamLeadID");
            entity.Property(e => e.TeamName)
                .HasMaxLength(50)
                .HasColumnName("teamName");
            entity.Property(e => e.TeamStatusId).HasColumnName("teamStatusID");

            entity.HasOne(d => d.Categories).WithMany(p => p.TfaTeams)
                .HasForeignKey(d => d.CategoriesId)
                .OnDelete(DeleteBehavior.Cascade)
                .HasConstraintName("FK_TFA_TEAMS_TFA_CATEGORIES");

            entity.HasOne(d => d.TeamLead).WithMany(p => p.TfaTeams)
                .HasForeignKey(d => d.TeamLeadId)
                .HasConstraintName("FK_teamLead");

            entity.HasOne(d => d.TeamStatus).WithMany(p => p.TfaTeams)
                .HasForeignKey(d => d.TeamStatusId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_teamStatusID");

            entity.HasMany(d => d.CategoriesNavigation).WithMany(p => p.Teams)
                .UsingEntity<Dictionary<string, object>>(
                    "TfaTeamsCategory",
                    r => r.HasOne<TfaCategory>().WithMany()
                        .HasForeignKey("CategoriesId")
                        .OnDelete(DeleteBehavior.ClientSetNull)
                        .HasConstraintName("FK_categoriesID"),
                    l => l.HasOne<TfaTeam>().WithMany()
                        .HasForeignKey("TeamId")
                        .OnDelete(DeleteBehavior.ClientSetNull)
                        .HasConstraintName("FK_teamID"),
                    j =>
                    {
                        j.HasKey("TeamId", "CategoriesId").HasName("PK__TFA_TEAM__CBFDC45B14D2BAED");
                        j.ToTable("TFA_TEAMS_CATEGORIES");
                        j.IndexerProperty<int>("TeamId").HasColumnName("teamID");
                        j.IndexerProperty<int>("CategoriesId").HasColumnName("categoriesID");
                    });
        });

        modelBuilder.Entity<TfaTeamstatus>(entity =>
        {
            entity.HasKey(e => e.StatusId).HasName("PK__TFA_TEAM__36257A388E96798A");

            entity.ToTable("TFA_TEAMSTATUS");

            entity.Property(e => e.StatusId).HasColumnName("statusID");
            entity.Property(e => e.StatusName)
                .HasMaxLength(50)
                .HasColumnName("statusName");
        });

        modelBuilder.Entity<TfaUser>(entity =>
        {
            entity.HasKey(e => e.UsersId).HasName("PK__TFA_USER__788FDD2552307940");

            entity.ToTable("TFA_USERS");

            entity.Property(e => e.UsersId).HasColumnName("usersID");
            entity.Property(e => e.Contrasenia)
                .HasMaxLength(999)
                .IsUnicode(false)
                .HasColumnName("contrasenia");
            entity.Property(e => e.RolId).HasColumnName("rolID");
            entity.Property(e => e.RolIdaddional).HasColumnName("rolIDAddional");
            entity.Property(e => e.UserEmail)
                .HasMaxLength(100)
                .HasColumnName("userEmail");
            entity.Property(e => e.UserLastName)
                .HasMaxLength(50)
                .HasColumnName("userLastName");
            entity.Property(e => e.UserName)
                .HasMaxLength(50)
                .HasColumnName("userName");
            entity.Property(e => e.UserPoints).HasColumnName("userPoints");

            entity.HasOne(d => d.Rol).WithMany(p => p.TfaUserRols)
                .HasForeignKey(d => d.RolId)
                .HasConstraintName("FK_userRol");

            entity.HasOne(d => d.RolIdaddionalNavigation).WithMany(p => p.TfaUserRolIdaddionalNavigations)
                .HasForeignKey(d => d.RolIdaddional)
                .HasConstraintName("FK_userRolAddional");

            entity.HasMany(d => d.ColaboratorTeams).WithMany(p => p.ColaboratorUsers)
                .UsingEntity<Dictionary<string, object>>(
                    "TfaTeamsColaborator",
                    r => r.HasOne<TfaTeam>().WithMany()
                        .HasForeignKey("ColaboratorTeamId")
                        .OnDelete(DeleteBehavior.ClientSetNull)
                        .HasConstraintName("FK_colaboratorTeamID"),
                    l => l.HasOne<TfaUser>().WithMany()
                        .HasForeignKey("ColaboratorUsersId")
                        .OnDelete(DeleteBehavior.ClientSetNull)
                        .HasConstraintName("FK_colaboratorUsersID"),
                    j =>
                    {
                        j.HasKey("ColaboratorUsersId", "ColaboratorTeamId").HasName("PK__TFA_TEAM__C0A7A83EBB8D9BE9");
                        j.ToTable("TFA_TEAMS_COLABORATORS");
                        j.IndexerProperty<int>("ColaboratorUsersId").HasColumnName("colaboratorUsersID");
                        j.IndexerProperty<int>("ColaboratorTeamId").HasColumnName("colaboratorTeamID");
                    });
        });

        modelBuilder.Entity<TfaUserPoint>(entity =>
        {
            entity.HasKey(e => e.UserPointsId).HasName("PK__TFA_USER__9208B819F7BA6D7A");

            entity.ToTable("TFA_USER_POINTS");

            entity.Property(e => e.UserPointsId).HasColumnName("UserPointsID");
            entity.Property(e => e.UserId).HasColumnName("UserID");

            entity.HasOne(d => d.User).WithMany(p => p.TfaUserPoints)
                .HasForeignKey(d => d.UserId)
                .OnDelete(DeleteBehavior.Cascade)
                .HasConstraintName("FK_TFA_USER_POINTS_User");
        });

        modelBuilder.Entity<TfaUsersTask>(entity =>
        {
            entity.HasKey(e => new { e.UserId, e.UserTaskId }).HasName("PK__TFA_USER__24DCCC70C6F17F8C");

            entity.ToTable("TFA_USERS_TASKS");

            entity.Property(e => e.UserId).HasColumnName("userID");
            entity.Property(e => e.UserTaskId).HasColumnName("userTaskID");
            entity.Property(e => e.EvidencePath)
                .HasMaxLength(255)
                .HasColumnName("evidencePath");
            entity.Property(e => e.EvidenceUploadDate)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("evidenceUploadDate");
            entity.Property(e => e.StatusTask)
                .HasDefaultValue(false)
                .HasColumnName("statusTask");

            entity.HasOne(d => d.User).WithMany(p => p.TfaUsersTasks)
                .HasForeignKey(d => d.UserId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_userID");

            entity.HasOne(d => d.UserTask).WithMany(p => p.TfaUsersTasks)
                .HasForeignKey(d => d.UserTaskId)
                .OnDelete(DeleteBehavior.ClientSetNull)
                .HasConstraintName("FK_userTasksID");
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
