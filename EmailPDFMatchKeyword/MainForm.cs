
namespace EmailPDFMatchKeyword
{
  public partial class MainForm : Form
  {
    private ExtractMethod _ExtractMethod;

    public MainForm()
    {
      InitializeComponent();
      Text = "Email Attachment Watcher - Bill to peer";
      Width = 850;
      Height = 600;
      InitUI();
      _ExtractMethod = new ExtractMethod(this);
    }
  }
}
