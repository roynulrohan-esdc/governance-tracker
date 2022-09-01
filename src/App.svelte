<script lang="ts">
  import { parseExcel, loading, readyToExport, exportPowerpoint } from "./stores/cad";

  let files: FileList;
</script>

<main>
  <div class="container">
    <div class="mrgn-tp-lg">
      <h1 id="title">Project Progress Tool</h1>
    </div>

    <div class="mrgn-tp-lg flex">
      <div>
        <div class="form-group">
          <label for="fileInput">Upload excel</label>
          <input id="fileInput" type="file" bind:files disabled={$loading} />
        </div>
        {#if files && files[0]}
          <button
            class="btn btn-primary mrgn-tp-md"
            disabled={$loading}
            on:click={() => {
              parseExcel(files[0]);
            }}
          >
            Convert
          </button>
        {/if}
      </div>
      <div>
        {#if $readyToExport}
          <div class="form-group">
            <label for="exportFile">Export powerpoint</label>
            <p>{files[0].name.split(".")[0]}.pptx</p>
            <button
              id="exportFile"
              class="btn btn-success mrgn-tp-md"
              on:click={() => {
                exportPowerpoint(files[0].name.split(".")[0]);
              }}>Download</button
            >
          </div>
        {/if}
      </div>
    </div>
  </div>
</main>

<style>
  main {
    padding-bottom: 20px;
    min-height: 100vh;
  }

  #title {
    border-bottom: 3px solid #af3c43;
    margin-bottom: 0.2em;
    margin-top: 1em;
    padding-bottom: 0.2em;
  }
</style>
